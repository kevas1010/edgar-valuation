"""Pure-Python DCF that mirrors the workbook formulas line for line.

The workbook is the product; this module is the reference implementation the
tests compare it against (and what the CLI uses to print a quick summary).
"""
from __future__ import annotations

from dataclasses import dataclass


def cost_of_equity(risk_free: float, beta: float, equity_risk_premium: float) -> float:
    """CAPM."""
    return risk_free + beta * equity_risk_premium


def wacc(risk_free: float, beta: float, equity_risk_premium: float, cost_of_debt: float, tax_rate: float,
         market_cap: float, debt: float, debt_materiality: float) -> float:
    """Weighted average cost of capital. Debt below the materiality threshold is ignored (all-equity)."""
    debt_share = debt / (market_cap + debt) if (market_cap + debt) else 0.0
    w_debt = debt_share if debt_share >= debt_materiality else 0.0
    return (1 - w_debt) * cost_of_equity(risk_free, beta, equity_risk_premium) + w_debt * cost_of_debt * (1 - tax_rate)


@dataclass
class DCFResult:
    revenue: list[float]
    ufcf: list[float]
    pv_ufcf: list[float]
    terminal_fcf: float
    terminal_value: float
    enterprise_value: float
    equity_value: float
    value_per_share: float


def dcf(base_revenue: float, growth: list[float], ebitda_margin: float, da_pct: float, capex_pct: float,
        tax_rate: float, discount_rate: float, terminal_growth: float, debt: float, cash: float, shares: float,
        terminal_capex_equals_da: bool = True) -> DCFResult:
    revenue, ufcf, pv = [], [], []
    rev = base_revenue
    for t, g in enumerate(growth, start=1):
        rev *= 1 + g
        da = rev * da_pct
        ebit = rev * ebitda_margin - da
        nopat = ebit - max(ebit, 0.0) * tax_rate          # no tax credit on losses
        fcf = nopat + da - rev * capex_pct
        revenue.append(rev)
        ufcf.append(fcf)
        pv.append(fcf / (1 + discount_rate) ** t)

    n = len(growth)
    terminal_capex = da if terminal_capex_equals_da else rev * capex_pct
    terminal_fcf = nopat + da - terminal_capex
    tv = terminal_fcf * (1 + terminal_growth) / (discount_rate - terminal_growth)   # Gordon growth
    ev = sum(pv) + tv / (1 + discount_rate) ** n
    equity = ev - debt + cash
    return DCFResult(revenue, ufcf, pv, terminal_fcf, tv, ev, equity, equity / shares)
