"""Comparable-companies analysis: peer trading multiples from EDGAR + yfinance.

    Market cap = price × shares
    EV         = market cap + debt − cash
    EV/Revenue, EV/EBITDA, P/E  (a multiple of a non-positive metric is meaningless -> excluded)

The workbook recomputes all of these with formulas; this module gathers the raw
inputs and records which peers are excluded from which multiple.
"""
from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
import requests

from edgar.client import EdgarClient
from edgar.extract import DataNote, extract_financials, with_derived
from model.prices import get_prices

MULTIPLES = {"EV/Revenue": "revenue", "EV/EBITDA": "ebitda", "P/E": "net_income"}
METRIC_LABELS = {"revenue": "revenue", "ebitda": "EBITDA", "net_income": "net income"}
# Above this, the denominator is usually a trough year and the multiple says nothing about value.
# Treated as "n/m" (same as a negative metric); editable on the Comps sheet.
MAX_MULTIPLE = 100.0


@dataclass
class PeerData:
    ticker: str
    name: str
    fiscal_year: int
    price: float
    shares: float
    debt: float
    cash: float
    revenue: float
    ebitda: float
    net_income: float

    @property
    def market_cap(self) -> float:
        return self.price * self.shares

    @property
    def enterprise_value(self) -> float:
        return self.market_cap + self.debt - self.cash

    def raw_multiple(self, name: str) -> float:
        """NaN when the denominator is missing or not positive."""
        metric = getattr(self, MULTIPLES[name])
        if not metric > 0:   # also catches NaN
            return float("nan")
        numerator = self.market_cap if name == "P/E" else self.enterprise_value
        return numerator / metric

    def multiple(self, name: str, cap: float = MAX_MULTIPLE) -> float:
        """The multiple used in the median: NaN ("n/m") if not meaningful or above the outlier cap."""
        m = self.raw_multiple(name)
        return m if m <= cap else float("nan")


def load_peer(client: EdgarClient, ticker: str, notes: list[DataNote], price_loader=get_prices) -> PeerData | None:
    ticker = ticker.strip().upper()
    try:
        fin = extract_financials(client.company_facts(ticker), ticker)
        price = price_loader(ticker)["current_price"]
    except (ValueError, KeyError, FileNotFoundError, requests.RequestException) as exc:
        notes.append(DataNote("Peer", ticker, f"Skipped entirely: {exc}"))
        return None
    if np.isnan(fin.shares_outstanding):
        notes.append(DataNote("Peer", ticker, "Skipped entirely: no share count, so no market cap."))
        return None

    t = with_derived(fin.table)
    fy = int(t.columns[-1])
    latest = lambda item: float(t.loc[item, fy])   # NaN if missing
    cash = latest("Cash")
    if np.isnan(cash):
        notes.append(DataNote("Peer", ticker, f"FY{fy} cash missing; treated as 0 in EV."))
        cash = 0.0
    peer = PeerData(
        ticker=ticker, name=fin.entity_name, fiscal_year=fy, price=float(price),
        shares=fin.shares_outstanding, debt=latest("Debt"), cash=cash,
        revenue=latest("Revenue"), ebitda=latest("EBITDA"), net_income=latest("Net Income"),
    )
    for name, attr in MULTIPLES.items():
        value = getattr(peer, attr)
        if not value > 0:
            why = "missing" if np.isnan(value) else f"negative (-${abs(value) / 1e6:,.1f}M)"
            notes.append(DataNote("Peer", ticker, f"FY{fy} {METRIC_LABELS[attr]} {why}; excluded from {name}."))
        elif peer.raw_multiple(name) > MAX_MULTIPLE:
            notes.append(DataNote("Peer", ticker, f"{name} of {peer.raw_multiple(name):,.0f}x exceeds the "
                                  f"{MAX_MULTIPLE:.0f}x outlier cap; excluded from {name} as n/m "
                                  f"(FY{fy} {METRIC_LABELS[attr]} ${value / 1e6:,.1f}M)."))
    return peer


def load_peers(client: EdgarClient, tickers: list[str], notes: list[DataNote], price_loader=get_prices) -> list[PeerData]:
    peers = [load_peer(client, t, notes, price_loader) for t in tickers if t.strip()]
    return [p for p in peers if p is not None]


def implied_value_per_share(peers: list[PeerData], revenue: float, ebitda: float, net_income: float,
                            debt: float, cash: float, shares: float) -> dict[str, float]:
    """Apply peer median multiples to the target's metrics (mirrors the Comps sheet).
    EV multiples give EV -> subtract debt, add cash; P/E gives equity value directly."""
    target = {"EV/Revenue": revenue, "EV/EBITDA": ebitda, "P/E": net_income}
    out = {}
    for name, metric in target.items():
        valid = [m for m in (p.multiple(name) for p in peers) if not np.isnan(m)]
        median = float(np.median(valid)) if valid else float("nan")
        if np.isnan(median) or not metric > 0:
            out[name] = float("nan")
            continue
        equity = metric * median if name == "P/E" else metric * median - debt + cash
        out[name] = equity / shares
    meaningful = [v for v in out.values() if not np.isnan(v)]
    out["Average"] = float(np.mean(meaningful)) if meaningful else float("nan")
    return out


def multiples_table(peers: list[PeerData]) -> pd.DataFrame:
    """Peer multiples plus median / mean rows (NaN = excluded), for the console and tests."""
    df = pd.DataFrame({name: [p.multiple(name) for p in peers] for name in MULTIPLES},
                      index=[p.ticker for p in peers])
    df.loc["Median"] = df.median()
    df.loc["Mean"] = df.iloc[:-1].mean()
    return df
