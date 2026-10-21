"""The end-to-end pipeline shared by the single-ticker CLI and portfolio mode:
pull data -> assemble the model inputs -> (optionally) re-run the model in Python."""
from __future__ import annotations

from datetime import date, datetime
from typing import Callable

import pandas as pd

from edgar.client import EdgarClient
from edgar.extract import extract_financials, with_derived
from excel.writer import ModelData
from model.assumptions import default_assumptions, default_scenarios
from model.beta import regress_beta
from model.comps import implied_value_per_share, load_peers, multiples_table
from model.dcf import dcf, wacc
from model.prices import MARKET_INDEX, get_prices, weekly_series

PCT_ROWS = {"Revenue Growth %", "EBITDA Margin %", "FCF Margin %", "Effective Tax Rate %"}


def format_historicals(table: pd.DataFrame) -> pd.DataFrame:
    """$ millions with 1 decimal, percentages with 1 decimal, blanks for missing."""
    def fmt(row_name, v):
        if pd.isna(v):
            return ""
        return f"{v:.1%}" if row_name in PCT_ROWS else f"{v / 1e6:,.1f}"
    out = table.copy().astype(object)
    for row in out.index:
        out.loc[row] = [fmt(row, v) for v in table.loc[row]]
    out.columns = [f"FY{c}" for c in out.columns]
    return out


def build_model(ticker: str, peer_tickers: list[str], client: EdgarClient,
                price_loader: Callable[[str], dict] = get_prices, as_of: date | None = None,
                pulled_at: datetime | None = None, verbose: bool = False) -> ModelData:
    """Pull every input and assemble the model. Data sources are injectable, so tests run offline."""
    say = print if verbose else (lambda *a, **k: None)

    say(f"Pulling SEC EDGAR financials for {ticker.upper()}...")
    fin = extract_financials(client.company_facts(ticker), ticker)
    say(format_historicals(with_derived(fin.table)).to_string())

    say(f"\nPulling prices and running beta regression vs {MARKET_INDEX}...")
    stock_px, market_px = price_loader(fin.ticker), price_loader(MARKET_INDEX)
    beta = regress_beta(weekly_series(stock_px), weekly_series(market_px))
    say(f"  Price ${stock_px['current_price']:,.2f} | beta {beta.beta:.2f} (R2 {beta.r_squared:.2f}, n={beta.n_obs})")

    peers = []
    if peer_tickers:
        say(f"\nPulling peers: {', '.join(peer_tickers)}...")
        peers = load_peers(client, peer_tickers, fin.notes, price_loader)
        if peers:
            say(multiples_table(peers).map(lambda v: "n/m" if pd.isna(v) else f"{v:.1f}x").to_string())

    assumptions = default_assumptions(fin, beta.beta, fin.notes)
    return ModelData(fin=fin, assumptions=assumptions, scenarios=default_scenarios(assumptions), beta=beta,
                     price=stock_px["current_price"], price_date=stock_px["price_date"],
                     as_of=as_of or date.today(), peers=peers, peer_tickers=peer_tickers,
                     pulled_at=pulled_at or datetime.now())


def python_values(data: ModelData, terminal_growth: float | None = None) -> dict[str, float]:
    """Re-run the model in Python (mirrors the workbook formulas). Used for console output, portfolio
    mode, and the tests that check the workbook against it."""
    fin, a = data.fin, data.assumptions
    g = a.terminal_growth if terminal_growth is None else terminal_growth
    debt, cash = fin.latest("Debt"), fin.latest("Cash")
    r = wacc(a.risk_free, a.beta, a.equity_risk_premium, a.cost_of_debt, a.tax_rate,
             data.price * fin.shares_outstanding, debt, a.debt_materiality)
    out = {"WACC": r}
    for s in data.scenarios:
        growth = a.revenue_growth if s.name == "Base" else s.revenue_growth       # Base links to Assumptions
        margin = a.ebitda_margin if s.name == "Base" else s.ebitda_margin
        result = dcf(fin.latest("Revenue"), growth, margin, a.da_pct, a.capex_pct, a.tax_rate, r, g,
                     debt, cash, fin.shares_outstanding, a.terminal_capex_equals_da)
        out[s.name], out[f"{s.name}EV"] = result.value_per_share, result.enterprise_value
    out["Target"] = sum(s.probability * out[s.name] for s in data.scenarios) / sum(s.probability for s in data.scenarios)
    t = with_derived(fin.table)
    out["Comps"] = implied_value_per_share(data.peers, fin.latest("Revenue"), t.loc["EBITDA"].iloc[-1],
                                           fin.latest("Net Income"), debt, cash, fin.shares_outstanding)["Average"]
    return out
