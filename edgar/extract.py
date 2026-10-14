"""Turn a raw SEC companyfacts JSON into a tidy 5-year table of annual financials.

Rules (see README for the reasoning):
  * Annual values only: form == "10-K" and fp == "FY".
  * 10-Ks also contain quarterly and prior-period comparatives, so duration items
    (income / cash-flow statement) must span ~1 year, and every value is matched
    to a fiscal-year-end date.
  * Restatements: the same period appears in several filings; keep the latest `filed`.
  * Each line item has an ordered list of fallback tags, resolved *per year*
    because companies switch tags over time.
  * Nothing missing ever raises: the value is NaN and a DataNote is recorded.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

ANNUAL_DAYS = (350, 380)   # a 52/53-week fiscal year is 364 or 371 days
YEAR_END_TOLERANCE_DAYS = 7


@dataclass(frozen=True)
class LineItem:
    name: str
    tags: tuple[str, ...]
    duration: bool  # True = flow over the year (P&L, cash flow); False = point in time (balance sheet)


LINE_ITEMS: tuple[LineItem, ...] = (
    LineItem("Revenue", ("RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet",
                         "RevenueFromContractWithCustomerIncludingAssessedTax"), True),
    LineItem("Operating Income", ("OperatingIncomeLoss",), True),
    LineItem("Net Income", ("NetIncomeLoss",), True),
    LineItem("D&A", ("DepreciationDepletionAndAmortization", "DepreciationAndAmortization", "Depreciation"), True),
    LineItem("Operating Cash Flow", ("NetCashProvidedByUsedInOperatingActivities",), True),
    LineItem("CapEx", ("PaymentsToAcquirePropertyPlantAndEquipment",), True),
    LineItem("Pretax Income", ("IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",), True),
    LineItem("Income Tax", ("IncomeTaxExpenseBenefit",), True),
    LineItem("Cash", ("CashAndCashEquivalentsAtCarryingValue",), False),
)
DEBT_TOTAL_TAG = "LongTermDebt"
DEBT_PARTS_TAGS = ("LongTermDebtNoncurrent", "LongTermDebtCurrent")


@dataclass
class DataNote:
    category: str   # "Missing", "Fallback", "Assumption", "Peer", "Info"
    item: str
    message: str


@dataclass
class Financials:
    ticker: str
    entity_name: str
    table: pd.DataFrame                  # rows = line items, columns = fiscal years (int), values in USD
    sources: dict[str, str]              # line item -> XBRL tag(s) actually used
    shares_outstanding: float
    shares_as_of: str
    notes: list[DataNote] = field(default_factory=list)
    year_ends: dict[int, pd.Timestamp] = field(default_factory=dict)  # fiscal year -> period-end date

    def latest(self, item: str) -> float:
        return self.table.loc[item].iloc[-1]


# ---------------------------------------------------------------- low-level helpers

def fiscal_year(end: pd.Timestamp) -> int:
    """Label a fiscal year by the calendar year it ends in. A 52/53-week year that
    spills a few days into January belongs to the prior year."""
    return end.year - 1 if (end.month == 1 and end.day <= YEAR_END_TOLERANCE_DAYS) else end.year


def annual_facts(facts: dict, tag: str, duration: bool, taxonomy: str = "us-gaap", unit: str = "USD") -> pd.Series:
    """All 10-K/FY values for one tag, deduped by period end (latest filing wins).

    Returns a Series indexed by period-end Timestamp, sorted ascending. Empty if the tag is absent.
    """
    entries = facts.get("facts", {}).get(taxonomy, {}).get(tag, {}).get("units", {}).get(unit, [])
    df = pd.DataFrame(entries)
    if df.empty:
        return pd.Series(dtype=float)

    df = df[(df["form"] == "10-K") & (df["fp"] == "FY")].copy()
    df["end"] = pd.to_datetime(df["end"])
    if duration:
        if "start" not in df:
            return pd.Series(dtype=float)
        days = (df["end"] - pd.to_datetime(df["start"])).dt.days
        df = df[days.between(*ANNUAL_DAYS)]

    # Amended / restated values: same period end reported again in a later filing.
    df = df.sort_values("filed").drop_duplicates("end", keep="last")
    return df.set_index("end")["val"].astype(float).sort_index()


def _match_year_ends(series: pd.Series, year_ends: dict[int, pd.Timestamp]) -> dict[int, float]:
    """Map a fact series onto fiscal years, keeping only values dated at a fiscal-year end."""
    out: dict[int, float] = {}
    if series.empty:
        return out
    for fy, fy_end in year_ends.items():
        close = series[(series.index - fy_end).map(abs) <= pd.Timedelta(days=YEAR_END_TOLERANCE_DAYS)]
        if not close.empty:
            out[fy] = close.iloc[-1]
    return out


def _describe_sources(used: dict[int, str]) -> str:
    """'TagA' if one tag covers every year, else 'TagA (FY2022, FY2023); TagB (FY2024)'."""
    if not used:
        return ""
    by_tag: dict[str, list[int]] = {}
    for fy, tag in sorted(used.items()):
        by_tag.setdefault(tag, []).append(fy)
    if len(by_tag) == 1:
        return next(iter(by_tag))
    return "; ".join(f"{tag} ({', '.join(f'FY{y}' for y in yrs)})" for tag, yrs in by_tag.items())


# ---------------------------------------------------------------- main entry points

def fiscal_year_ends(facts: dict, n_years: int) -> dict[int, pd.Timestamp]:
    """The last `n_years` fiscal-year-end dates, discovered from revenue and net income."""
    ends: dict[int, pd.Timestamp] = {}
    anchor_tags = LINE_ITEMS[0].tags + ("NetIncomeLoss",)
    for tag in anchor_tags:
        for end in annual_facts(facts, tag, duration=True).index:
            ends.setdefault(fiscal_year(end), end)
    years = sorted(ends)[-n_years:]
    return {fy: ends[fy] for fy in years}


def extract_financials(facts: dict, ticker: str, n_years: int = 5) -> Financials:
    notes: list[DataNote] = []
    year_ends = fiscal_year_ends(facts, n_years)
    years = list(year_ends)
    table = pd.DataFrame(index=[i.name for i in LINE_ITEMS] + ["Debt"], columns=years, dtype=float)
    sources: dict[str, str] = {}

    for item in LINE_ITEMS:
        used: dict[int, str] = {}
        by_tag = {tag: _match_year_ends(annual_facts(facts, tag, item.duration), year_ends) for tag in item.tags}
        for fy in years:
            for tag in item.tags:  # first tag (in priority order) that has this year wins
                if fy in by_tag[tag]:
                    table.loc[item.name, fy] = by_tag[tag][fy]
                    used[fy] = tag
                    break
        sources[item.name] = _describe_sources(used)

        missing = [fy for fy in years if fy not in used]
        if missing:
            notes.append(DataNote("Missing", item.name,
                                  f"No value under any of {', '.join(item.tags)} for "
                                  f"{', '.join(f'FY{y}' for y in missing)}; left blank."))
        fallback_years = [fy for fy, tag in used.items() if tag != item.tags[0]]
        if fallback_years:
            notes.append(DataNote("Fallback", item.name,
                                  f"Primary tag {item.tags[0]} unavailable for "
                                  f"{', '.join(f'FY{y}' for y in fallback_years)}; used {sources[item.name]}."))

    sources["Debt"] = _extract_debt(facts, table, year_ends, notes)
    shares, as_of, shares_tag = _extract_shares(facts, notes)
    sources["Shares Outstanding"] = shares_tag

    return Financials(
        ticker=ticker.upper(),
        entity_name=facts.get("entityName", ticker.upper()),
        table=table,
        sources=sources,
        shares_outstanding=shares,
        shares_as_of=as_of,
        notes=notes,
        year_ends=year_ends,
    )


def _extract_debt(facts: dict, table: pd.DataFrame, year_ends: dict[int, pd.Timestamp],
                  notes: list[DataNote]) -> str:
    """Debt = LongTermDebt, else LongTermDebtNoncurrent + LongTermDebtCurrent, else 0."""
    total = _match_year_ends(annual_facts(facts, DEBT_TOTAL_TAG, duration=False), year_ends)
    parts = {t: _match_year_ends(annual_facts(facts, t, duration=False), year_ends) for t in DEBT_PARTS_TAGS}
    used: dict[int, str] = {}
    zero_years = []
    for fy in year_ends:
        if fy in total:
            table.loc["Debt", fy] = total[fy]
            used[fy] = DEBT_TOTAL_TAG
        elif any(fy in p for p in parts.values()):
            present = [t for t in DEBT_PARTS_TAGS if fy in parts[t]]
            table.loc["Debt", fy] = sum(parts[t][fy] for t in present)
            used[fy] = " + ".join(present)
        else:
            table.loc["Debt", fy] = 0.0
            used[fy] = "none (assumed 0)"
            zero_years.append(fy)
    if zero_years:
        notes.append(DataNote("Assumption", "Debt",
                              f"No long-term debt tags reported for {', '.join(f'FY{y}' for y in zero_years)}; "
                              "debt assumed to be 0."))
    return _describe_sources(used)


def _extract_shares(facts: dict, notes: list[DataNote]) -> tuple[float, str, str]:
    """Most recent cover-page share count. Multi-class companies report one entry per
    class in the same filing, so entries from the latest filing are summed."""
    entries = facts.get("facts", {}).get("dei", {}).get("EntityCommonStockSharesOutstanding", {}) \
                   .get("units", {}).get("shares", [])
    tag = "dei:EntityCommonStockSharesOutstanding"
    if not entries:
        # Fallback: balance-sheet share count from the latest 10-K.
        entries = [e for e in facts.get("facts", {}).get("us-gaap", {}).get("CommonStockSharesOutstanding", {})
                   .get("units", {}).get("shares", []) if e.get("form") == "10-K"]
        tag = "us-gaap:CommonStockSharesOutstanding"
        if entries:
            notes.append(DataNote("Fallback", "Shares Outstanding",
                                  "dei:EntityCommonStockSharesOutstanding missing; used balance-sheet share count."))
    if not entries:
        notes.append(DataNote("Missing", "Shares Outstanding", "No share count reported; left blank."))
        return float("nan"), "", ""

    df = pd.DataFrame(entries)
    latest = df[df["filed"] == df["filed"].max()]
    latest = latest[latest["end"] == latest["end"].max()]
    return float(latest["val"].sum()), str(latest["end"].iloc[0]), tag


def with_derived(table: pd.DataFrame) -> pd.DataFrame:
    """Add the derived lines used throughout the model. (The workbook rebuilds these as formulas.)"""
    t = table.copy()
    t.loc["EBITDA"] = t.loc["Operating Income"] + t.loc["D&A"]
    t.loc["Free Cash Flow"] = t.loc["Operating Cash Flow"] - t.loc["CapEx"]
    t.loc["Revenue Growth %"] = t.loc["Revenue"].pct_change()
    t.loc["EBITDA Margin %"] = t.loc["EBITDA"] / t.loc["Revenue"]
    t.loc["FCF Margin %"] = t.loc["Free Cash Flow"] / t.loc["Revenue"]
    t.loc["Effective Tax Rate %"] = t.loc["Income Tax"] / t.loc["Pretax Income"]
    return t
