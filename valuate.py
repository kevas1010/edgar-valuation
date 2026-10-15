"""CLI entry point:  python valuate.py AEHR --peers TER,COHU,FORM,KLIC"""
from __future__ import annotations

import argparse
from datetime import date, datetime
from pathlib import Path

import pandas as pd

from edgar.client import EdgarClient
from edgar.extract import extract_financials, with_derived
from excel.writer import ModelData, write_workbook
from model.assumptions import default_assumptions, default_scenarios
from model.beta import regress_beta
from model.comps import implied_value_per_share, load_peers, multiples_table
from model.dcf import dcf, wacc
from model.prices import MARKET_INDEX, get_prices, weekly_series

PCT_ROWS = {"Revenue Growth %", "EBITDA Margin %", "FCF Margin %", "Effective Tax Rate %"}
OUTPUT_DIR = Path(__file__).resolve().parent / "output"


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


def python_values(data: ModelData) -> dict[str, float]:
    """Re-run the model in Python (mirrors the workbook) for the console summary."""
    fin, a = data.fin, data.assumptions
    debt, cash = fin.latest("Debt"), fin.latest("Cash")
    r = wacc(a.risk_free, a.beta, a.equity_risk_premium, a.cost_of_debt, a.tax_rate,
             data.price * fin.shares_outstanding, debt, a.debt_materiality)
    out = {"WACC": r}
    for s in data.scenarios:
        out[s.name] = dcf(fin.latest("Revenue"), s.revenue_growth, s.ebitda_margin, a.da_pct, a.capex_pct,
                          a.tax_rate, r, a.terminal_growth, debt, cash, fin.shares_outstanding,
                          a.terminal_capex_equals_da).value_per_share
    out["Target"] = sum(s.probability * out[s.name] for s in data.scenarios)
    t = with_derived(fin.table)
    out["Comps"] = implied_value_per_share(data.peers, fin.latest("Revenue"), t.loc["EBITDA"].iloc[-1],
                                           fin.latest("Net Income"), debt, cash, fin.shares_outstanding)["Average"]
    return out


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Build an Excel valuation model from SEC EDGAR filings.")
    p.add_argument("ticker", help="Target company ticker, e.g. AEHR")
    p.add_argument("--peers", default="", help="Comma-separated peer tickers, e.g. TER,COHU,FORM,KLIC")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    today, pulled_at = date.today(), datetime.now()
    peer_tickers = [t.strip().upper() for t in args.peers.split(",") if t.strip()]
    client = EdgarClient()

    print(f"Pulling SEC EDGAR financials for {args.ticker.upper()}...")
    fin = extract_financials(client.company_facts(args.ticker), args.ticker)
    with pd.option_context("display.width", 200, "display.max_columns", 20):
        print(format_historicals(with_derived(fin.table)).to_string())

    print(f"\nPulling prices and running beta regression vs {MARKET_INDEX}...")
    stock_px, market_px = get_prices(fin.ticker), get_prices(MARKET_INDEX)
    beta = regress_beta(weekly_series(stock_px), weekly_series(market_px))
    print(f"  Price ${stock_px['current_price']:,.2f} | beta {beta.beta:.2f} (R2 {beta.r_squared:.2f}, n={beta.n_obs})")

    peers = []
    if peer_tickers:
        print(f"\nPulling peers: {', '.join(peer_tickers)}...")
        peers = load_peers(client, peer_tickers, fin.notes)
        if peers:
            print(multiples_table(peers).map(lambda v: "n/m" if pd.isna(v) else f"{v:.1f}x").to_string())

    assumptions = default_assumptions(fin, beta.beta, fin.notes)
    data = ModelData(fin=fin, assumptions=assumptions, scenarios=default_scenarios(assumptions), beta=beta,
                     price=stock_px["current_price"], price_date=stock_px["price_date"], as_of=today,
                     peers=peers, peer_tickers=peer_tickers, pulled_at=pulled_at)
    path = write_workbook(data, OUTPUT_DIR / f"{fin.ticker}_valuation_{today.isoformat()}.xlsx")

    v = python_values(data)
    print(f"\nWACC {v['WACC']:.1%} | Bear ${v['Bear']:,.2f} | Base (DCF) ${v['Base']:,.2f} | Bull ${v['Bull']:,.2f}")
    print("Comps-implied value " + ("n/m" if pd.isna(v["Comps"]) else f"${v['Comps']:,.2f}"))
    print(f"Scenario-weighted target ${v['Target']:,.2f} vs price ${data.price:,.2f} "
          f"({v['Target'] / data.price - 1:+.1%})")
    print(f"\n{len(fin.notes)} data notes (see the Data Notes sheet)")
    print(f"Wrote {path}")


if __name__ == "__main__":
    main()
