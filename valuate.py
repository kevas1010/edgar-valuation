"""CLI entry point:  python valuate.py AEHR --peers TER,COHU,FORM,KLIC"""
from __future__ import annotations

import argparse

import pandas as pd

from edgar.client import EdgarClient
from edgar.extract import extract_financials, with_derived
from model.beta import regress_beta
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


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Build an Excel valuation model from SEC EDGAR filings.")
    p.add_argument("ticker", help="Target company ticker, e.g. AEHR")
    p.add_argument("--peers", default="", help="Comma-separated peer tickers, e.g. TER,COHU,FORM,KLIC")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    client = EdgarClient()
    fin = extract_financials(client.company_facts(args.ticker), args.ticker)

    print(f"\n{fin.entity_name} ({fin.ticker})  -  $ millions unless noted\n")
    with pd.option_context("display.width", 200, "display.max_columns", 20):
        print(format_historicals(with_derived(fin.table)).to_string())
    print(f"\nShares outstanding: {fin.shares_outstanding:,.0f} (as of {fin.shares_as_of})")

    print("\nSource tags:")
    for item, src in fin.sources.items():
        print(f"  {item:<20} {src}")
    if fin.notes:
        print("\nData notes:")
        for n in fin.notes:
            print(f"  [{n.category}] {n.item}: {n.message}")

    stock_px, market_px = get_prices(fin.ticker), get_prices(MARKET_INDEX)
    beta = regress_beta(weekly_series(stock_px), weekly_series(market_px))
    print(f"\nCurrent price: ${stock_px['current_price']:,.2f} (as of {stock_px['price_date']})")
    print(f"\nBeta regression: weekly returns vs {MARKET_INDEX}, "
          f"{beta.returns.index[0]:%Y-%m-%d} to {beta.returns.index[-1]:%Y-%m-%d}")
    print(f"  Beta            {beta.beta:8.3f}   (std err {beta.std_err:.3f}, "
          f"95% CI {beta.beta - 1.96 * beta.std_err:.2f} to {beta.beta + 1.96 * beta.std_err:.2f})")
    print(f"  Alpha (weekly)  {beta.alpha:8.4f}")
    print(f"  R-squared       {beta.r_squared:8.3f}")
    print(f"  Observations    {beta.n_obs:8d}")
    print(f"  Blume-adjusted  {beta.adjusted_beta:8.3f}   (reference only; model uses raw beta)")
    print("\n  Chart data (first 5 of {} points):".format(len(beta.returns)))
    print(beta.returns.head().map(lambda v: f"{v:+.2%}").to_string())
    print("\n  Fit line endpoints:")
    print(beta.fit_line().map(lambda v: f"{v:+.2%}").to_string(index=False))


if __name__ == "__main__":
    main()
