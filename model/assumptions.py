"""Default DCF assumptions derived from the company's own history.

These become the blue input cells on the Assumptions sheet. They are starting
points, not forecasts: every one can be overwritten in Excel.
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from edgar.extract import DataNote, Financials, with_derived

RISK_FREE = 0.0425
EQUITY_RISK_PREMIUM = 0.055
TERMINAL_GROWTH = 0.03
PRE_TAX_COST_OF_DEBT = 0.06
DEBT_MATERIALITY = 0.05          # debt below 5% of total capital -> treat as all-equity
TAX_CLAMP = (0.0, 0.30)
GROWTH_CLAMP = (-0.20, 0.50)     # keeps a single freak year from producing a silly Year-1 growth rate

# Used only when the history needed for a default is missing (always flagged in Data Notes).
PLACEHOLDERS = {"ebitda_margin": 0.15, "da_pct": 0.03, "capex_pct": 0.03, "tax_rate": 0.21, "growth": 0.03}


@dataclass
class Assumptions:
    revenue_growth: list[float]      # Year 1..5
    ebitda_margin: float
    da_pct: float
    capex_pct: float
    tax_rate: float
    beta: float
    risk_free: float = RISK_FREE
    equity_risk_premium: float = EQUITY_RISK_PREMIUM
    terminal_growth: float = TERMINAL_GROWTH
    cost_of_debt: float = PRE_TAX_COST_OF_DEBT
    debt_materiality: float = DEBT_MATERIALITY
    basis: dict[str, str] = field(default_factory=dict)   # how each default was derived (Assumptions notes column)


def _weighted_ratio(num: pd.Series, den: pd.Series) -> float:
    """sum(num) / sum(den) over years where both exist -> revenue-weighted average ratio."""
    both = num.notna() & den.notna()
    total = den[both].sum()
    return float(num[both].sum() / total) if both.any() and total != 0 else float("nan")


def revenue_cagr(revenue: pd.Series) -> float:
    rev = revenue.dropna()
    if len(rev) < 2 or rev.iloc[0] <= 0 or rev.iloc[-1] <= 0:
        return float("nan")
    years = rev.index[-1] - rev.index[0]
    return float((rev.iloc[-1] / rev.iloc[0]) ** (1 / years) - 1)


def fade(start: float, end: float, n: int = 5) -> list[float]:
    """Straight line from `start` (Year 1) to `end` (Year n)."""
    return [float(x) for x in np.linspace(start, end, n)]


def default_assumptions(fin: Financials, beta: float, notes: list[DataNote]) -> Assumptions:
    t = with_derived(fin.table)
    years = f"FY{t.columns[0]}–FY{t.columns[-1]}"

    def pick(key: str, value: float, label: str, how: str) -> tuple[float, str]:
        if np.isnan(value):
            fallback = PLACEHOLDERS[key]
            notes.append(DataNote("Assumption", label,
                                  f"Not enough history to compute a default; used placeholder {fallback:.1%}."))
            return fallback, f"Placeholder {fallback:.1%} (history unavailable) – please review"
        return value, how

    # Revenue growth: historical CAGR, faded to terminal growth by Year 5.
    cagr = revenue_cagr(t.loc["Revenue"])
    g1, g1_basis = pick("growth", cagr, "Revenue growth", f"{years} revenue CAGR ({cagr:.1%})")
    clamped = float(np.clip(g1, *GROWTH_CLAMP))
    if clamped != g1:
        notes.append(DataNote("Assumption", "Revenue growth",
                              f"Historical CAGR {g1:.1%} clamped to {clamped:.1%} for Year 1."))
        g1_basis += f", clamped to {clamped:.1%}"

    margin, margin_basis = pick("ebitda_margin", _weighted_ratio(t.loc["EBITDA"], t.loc["Revenue"]),
                                "EBITDA margin", f"{years} revenue-weighted average")
    da, da_basis = pick("da_pct", _weighted_ratio(t.loc["D&A"], t.loc["Revenue"]),
                        "D&A % of revenue", f"{years} revenue-weighted average")
    capex, capex_basis = pick("capex_pct", _weighted_ratio(t.loc["CapEx"], t.loc["Revenue"]),
                              "CapEx % of revenue", f"{years} revenue-weighted average")

    # Tax: effective rate over profitable years only (a tax rate on a loss is meaningless).
    profitable = t.loc["Pretax Income"] > 0
    raw_tax = _weighted_ratio(t.loc["Income Tax"][profitable], t.loc["Pretax Income"][profitable])
    tax, tax_basis = pick("tax_rate", raw_tax, "Tax rate", f"{years} effective rate, profitable years ({raw_tax:.1%})")
    clamped_tax = float(np.clip(tax, *TAX_CLAMP))
    if clamped_tax != tax:
        notes.append(DataNote("Assumption", "Tax rate",
                              f"Historical effective rate {tax:.1%} clamped to {clamped_tax:.1%} (allowed range 0-30%)."))
        tax_basis += f", clamped to {clamped_tax:.1%}"

    return Assumptions(
        revenue_growth=fade(clamped, TERMINAL_GROWTH),
        ebitda_margin=margin,
        da_pct=da,
        capex_pct=capex,
        tax_rate=clamped_tax,
        beta=beta,
        basis={
            "revenue_growth": f"Year 1 = {g1_basis}; fades linearly to terminal growth by Year 5",
            "ebitda_margin": margin_basis,
            "da_pct": da_basis,
            "capex_pct": capex_basis,
            "tax_rate": tax_basis,
            "beta": "OLS regression, 2y weekly returns vs S&P 500 (see Beta Regression sheet)",
            "risk_free": "Approx. 10-year US Treasury yield",
            "equity_risk_premium": "Long-run US equity risk premium",
            "terminal_growth": "Long-run nominal GDP-like growth",
            "cost_of_debt": "Pre-tax cost of debt (book debt only)",
            "debt_materiality": "Debt below this share of capital is ignored (all-equity WACC)",
        },
    )
