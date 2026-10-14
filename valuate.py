"""CLI entry point:  python valuate.py AEHR --peers TER,COHU,FORM,KLIC"""
from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from edgar.client import EdgarClient
from edgar.extract import extract_financials, with_derived
from excel.writer import ModelData, write_workbook
from model.assumptions import default_assumptions
from model.beta import regress_beta
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


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Build an Excel valuation model from SEC EDGAR filings.")
    p.add_argument("ticker", help="Target company ticker, e.g. AEHR")
    p.add_argument("--peers", default="", help="Comma-separated peer tickers, e.g. TER,COHU,FORM,KLIC")
    return p.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    today = date.today()
    client = EdgarClient()

    print(f"Pulling SEC EDGAR financials for {args.ticker.upper()}...")
    fin = extract_financials(client.company_facts(args.ticker), args.ticker)
    with pd.option_context("display.width", 200, "display.max_columns", 20):
        print(format_historicals(with_derived(fin.table)).to_string())

    print(f"Pulling prices and running beta regression vs {MARKET_INDEX}...")
    stock_px, market_px = get_prices(fin.ticker), get_prices(MARKET_INDEX)
    beta = regress_beta(weekly_series(stock_px), weekly_series(market_px))
    print(f"  Price ${stock_px['current_price']:,.2f} | beta {beta.beta:.2f} (R2 {beta.r_squared:.2f}, n={beta.n_obs})")

    assumptions = default_assumptions(fin, beta.beta, fin.notes)
    data = ModelData(fin=fin, assumptions=assumptions, beta=beta,
                     price=stock_px["current_price"], price_date=stock_px["price_date"], as_of=today)
    path = write_workbook(data, OUTPUT_DIR / f"{fin.ticker}_valuation_{today.isoformat()}.xlsx")

    for n in fin.notes:
        print(f"  [{n.category}] {n.item}: {n.message}")
    print(f"\nWrote {path}")


if __name__ == "__main__":
    main()
