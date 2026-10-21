"""Value every stock in a portfolio and track results over time.

Each stock gets its own full valuation workbook (the same model as the single-ticker CLI).
Cash, options, funds/ETFs and anything EDGAR can't value are kept for gain/loss tracking
but marked "Not modeled" with the reason. One bad holding never stops the run.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Callable

import pandas as pd
import requests

from edgar.client import EdgarClient
from excel.writer import write_workbook
from model.prices import get_prices
from portfolio.fidelity import CASH, OPTION, STOCK, Holding
from model.pipeline import build_model, python_values

MODELED, NOT_MODELED = "Modeled", "Not modeled"
SKIP_REASONS = {CASH: "Cash / money market", OPTION: "Option contract", "Other": "Unrecognized symbol format"}


@dataclass
class HoldingResult:
    holding: Holding
    status: str
    reason: str = ""
    dcf_value: float = math.nan
    comps_value: float = math.nan
    target: float = math.nan          # scenario-weighted price target
    workbook: Path | None = None

    @property
    def upside(self) -> float:
        """Model target vs the price in the Fidelity export. n/m (NaN) when the target is <= 0:
        a share can't be worth less than zero, so a negative model value is a warning, not a price."""
        if self.status != MODELED or not self.target > 0:
            return math.nan
        return self.target / self.holding.last_price - 1


def _why_not_modeled(exc: Exception) -> str:
    if isinstance(exc, requests.HTTPError) and exc.response is not None and exc.response.status_code == 404:
        return "No XBRL financials on EDGAR (fund, ETF, or foreign filer?)"
    if isinstance(exc, ValueError) and "not found in SEC company_tickers" in str(exc):
        return "Not an SEC-registered company ticker (mutual fund or ETF?)"
    if isinstance(exc, ValueError) and "no annual 10-K" in str(exc):
        return "No 10-K operating financials on EDGAR (ETF, fund or trust?)"
    return f"{type(exc).__name__}: {exc}"


def value_holding(h: Holding, client: EdgarClient, price_loader: Callable[[str], dict], as_of: date,
                  out_dir: Path, peers: list[str] | None = None) -> HoldingResult:
    if h.kind != STOCK:
        return HoldingResult(h, NOT_MODELED, SKIP_REASONS.get(h.kind, h.kind))
    try:
        data = build_model(h.sec_ticker, peers or [], client, price_loader, as_of=as_of)
        if math.isnan(data.fin.shares_outstanding):
            return HoldingResult(h, NOT_MODELED, "No share count reported on EDGAR")
        if math.isnan(data.fin.latest("Revenue")):
            return HoldingResult(h, NOT_MODELED, "No revenue or operating cash flows reported "
                                                 "(holding partnership, trust, or fund?)")
        path = write_workbook(data, out_dir / f"{data.fin.ticker}_valuation_{as_of.isoformat()}.xlsx")
        v = python_values(data)          # mirrors the workbook formulas (verified by tests/test_dcf.py)
    except Exception as exc:  # noqa: BLE001 - isolate each holding; the reason is reported, not swallowed
        return HoldingResult(h, NOT_MODELED, _why_not_modeled(exc))
    note = "" if v["Target"] > 0 else (
        "Model equity value is negative: on these assumptions projected cash flows don't cover "
        "the debt. Upside n/m; open the model and revisit margins/CapEx")
    return HoldingResult(h, MODELED, note, dcf_value=v["Base"], comps_value=v["Comps"], target=v["Target"],
                         workbook=path)


def value_portfolio(holdings: list[Holding], client: EdgarClient, out_dir: Path, as_of: date | None = None,
                    price_loader: Callable[[str], dict] = get_prices, peers: dict[str, list[str]] | None = None,
                    log: Callable[[str], None] = print) -> list[HoldingResult]:
    as_of = as_of or date.today()
    out_dir.mkdir(parents=True, exist_ok=True)
    results = []
    for h in holdings:
        log(f"  {h.symbol:<16} ...")
        r = value_holding(h, client, price_loader, as_of, out_dir, (peers or {}).get(h.symbol))
        log(f"  {h.symbol:<16} {r.status}" + (f" - {r.reason}" if r.reason else f" - target ${r.target:,.2f}"))
        results.append(r)
    return results


# ---------------------------------------------------------------- history (tracking over time)

HISTORY_COLUMNS = ["date", "symbol", "status", "quantity", "price", "market_value", "cost_basis", "gain",
                   "dcf_value", "comps_value", "target", "upside"]


def append_history(results: list[HoldingResult], history_path: Path, as_of: date) -> pd.DataFrame:
    """Add today's snapshot to the history CSV (replacing any earlier run from the same day)."""
    rows = [{
        "date": as_of.isoformat(), "symbol": r.holding.symbol, "status": r.status,
        "quantity": r.holding.quantity, "price": r.holding.last_price, "market_value": r.holding.current_value,
        "cost_basis": r.holding.cost_basis, "gain": r.holding.gain, "dcf_value": r.dcf_value,
        "comps_value": r.comps_value, "target": r.target, "upside": r.upside,
    } for r in results]
    new = pd.DataFrame(rows, columns=HISTORY_COLUMNS)
    if history_path.exists():
        old = pd.read_csv(history_path)
        new = pd.concat([old[old["date"] != as_of.isoformat()], new], ignore_index=True)
    history_path.parent.mkdir(parents=True, exist_ok=True)
    new.to_csv(history_path, index=False)
    return new


def history_totals(history: pd.DataFrame) -> pd.DataFrame:
    """One row per run date: portfolio value, gain, and value-weighted model upside of modeled holdings."""
    def summarize(day: pd.DataFrame) -> pd.Series:
        known = day[day["cost_basis"].notna()]
        modeled = day[day["upside"].notna()]          # modeled holdings with a meaningful (positive) target
        mv = modeled["market_value"].sum()
        return pd.Series({
            "market_value": day["market_value"].sum(),
            "cost_basis_known": known["cost_basis"].sum(),
            "gain": known["gain"].sum(),
            "gain_pct": known["gain"].sum() / known["cost_basis"].sum() if known["cost_basis"].sum() else math.nan,
            "modeled_value": mv,
            "weighted_upside": (modeled["market_value"] * modeled["upside"]).sum() / mv if mv else math.nan,
        })
    return history.groupby("date").apply(summarize, include_groups=False).reset_index()
