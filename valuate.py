"""CLI entry point.

    python valuate.py AEHR --peers TER,COHU,FORM,KLIC        # one company -> one valuation workbook
    python valuate.py --portfolio positions/Portfolio_Positions.csv   # every stock in a Fidelity export
"""
from __future__ import annotations

import argparse
import math
from datetime import date
from pathlib import Path

import pandas as pd

from edgar.client import EdgarClient
from excel.writer import write_workbook
from model.pipeline import build_model, python_values

OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def run_single(ticker: str, peers: str, out_dir: Path) -> None:
    peer_tickers = [t.strip().upper() for t in peers.split(",") if t.strip()]
    with pd.option_context("display.width", 200, "display.max_columns", 20):
        data = build_model(ticker, peer_tickers, EdgarClient(), verbose=True)
    path = write_workbook(data, out_dir / f"{data.fin.ticker}_valuation_{data.as_of.isoformat()}.xlsx")

    v = python_values(data)
    print(f"\nWACC {v['WACC']:.1%} | Bear ${v['Bear']:,.2f} | Base (DCF) ${v['Base']:,.2f} | Bull ${v['Bull']:,.2f}")
    print("Comps-implied value " + ("n/m" if pd.isna(v["Comps"]) else f"${v['Comps']:,.2f}"))
    print(f"Scenario-weighted target ${v['Target']:,.2f} vs price ${data.price:,.2f} "
          f"({v['Target'] / data.price - 1:+.1%})")
    print(f"\n{len(data.fin.notes)} data notes (see the Data Notes sheet)")
    print(f"Wrote {path}")


def run_portfolio(positions: Path, out_dir: Path) -> None:
    # Imported here so the single-ticker path doesn't depend on portfolio code.
    from excel.portfolio_writer import write_portfolio_workbook
    from portfolio.fidelity import parse_positions
    from portfolio.run import append_history, history_totals, value_portfolio

    today = date.today()
    holdings = parse_positions(positions)
    print(f"Read {len(holdings)} holdings from {positions.name}. Valuing each stock...")
    run_dir = out_dir / "portfolio" / today.isoformat()
    results = value_portfolio(holdings, EdgarClient(), run_dir, as_of=today)

    history = append_history(results, out_dir / "portfolio" / "history.csv", today)
    totals = history_totals(history)
    path = write_portfolio_workbook(results, history, totals, run_dir / f"Portfolio_{today.isoformat()}.xlsx", today)

    latest = totals.iloc[-1]
    upside = latest["weighted_upside"]
    print(f"\nPortfolio value ${latest['market_value']:,.2f} | gain ${latest['gain']:,.2f} ({latest['gain_pct']:+.1%}) "
          f"| modeled {sum(r.status == 'Modeled' for r in results)}/{len(results)} holdings"
          + ("" if math.isnan(upside) else f" | value-weighted model upside {upside:+.1%}"))
    print(f"Wrote {path}")


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Build Excel valuation models from SEC EDGAR filings.")
    p.add_argument("ticker", nargs="?", help="Target company ticker, e.g. AEHR")
    p.add_argument("--peers", default="", help="Comma-separated peer tickers, e.g. TER,COHU,FORM,KLIC")
    p.add_argument("--portfolio", type=Path, help="Fidelity positions CSV: value every stock in it")
    p.add_argument("--out-dir", type=Path, default=OUTPUT_DIR, help="Where to write workbooks (default: output/)")
    args = p.parse_args(argv)
    if bool(args.ticker) == bool(args.portfolio):
        p.error("give either a TICKER or --portfolio PATH (not both)")
    return args


def main(argv=None):
    args = parse_args(argv)
    if args.portfolio:
        run_portfolio(args.portfolio, args.out_dir)
    else:
        run_single(args.ticker, args.peers, args.out_dir)


if __name__ == "__main__":
    main()
