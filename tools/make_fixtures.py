"""Regenerate tests/fixtures/ from the local cache/ (run after a live `valuate.py` run).

Trims each SEC companyfacts JSON down to the tags the extractor reads, so the
fixtures are small enough to commit but are still *real* EDGAR data (including
restatements, quarterly values inside 10-Ks, and stale tags).

    python tools/make_fixtures.py --date 2026-10-14 AEHR TER COHU FORM KLIC
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from edgar.client import EdgarClient  # noqa: E402
from edgar.extract import DEBT_PARTS_TAGS, DEBT_TOTAL_TAG, LINE_ITEMS  # noqa: E402

CACHE, FIXTURES = ROOT / "cache", ROOT / "tests" / "fixtures"
KEEP_GAAP = {t for item in LINE_ITEMS for t in item.tags} | {DEBT_TOTAL_TAG, *DEBT_PARTS_TAGS,
                                                              "CommonStockSharesOutstanding"}
KEEP_DEI = {"EntityCommonStockSharesOutstanding"}


def trim(facts: dict) -> dict:
    out = {k: v for k, v in facts.items() if k != "facts"}
    out["facts"] = {
        "us-gaap": {k: v for k, v in facts["facts"].get("us-gaap", {}).items() if k in KEEP_GAAP},
        "dei": {k: v for k, v in facts["facts"].get("dei", {}).items() if k in KEEP_DEI},
    }
    return out


def main():
    p = argparse.ArgumentParser()
    p.add_argument("tickers", nargs="+")
    p.add_argument("--date", required=True, help="Price cache date to copy, e.g. 2026-10-14")
    args = p.parse_args()

    client = EdgarClient(cache_dir=CACHE, offline=True)
    tickers = [t.upper() for t in args.tickers]
    all_tickers = json.loads((CACHE / "company_tickers.json").read_text())
    subset = {k: v for k, v in all_tickers.items() if v["ticker"].upper() in tickers}
    (FIXTURES / "companyfacts").mkdir(parents=True, exist_ok=True)
    (FIXTURES / "prices").mkdir(parents=True, exist_ok=True)
    (FIXTURES / "company_tickers.json").write_text(json.dumps(subset, indent=1))

    for t in tickers:
        cik = client.ticker_to_cik(t)
        name = f"CIK{cik:010d}.json"
        facts = json.loads((CACHE / "companyfacts" / name).read_text())
        (FIXTURES / "companyfacts" / name).write_text(json.dumps(trim(facts)))
    for t in tickers + ["GSPC"]:
        shutil.copy(CACHE / "prices" / f"{t}_{args.date}.json", FIXTURES / "prices" / f"{t}_{args.date}.json")
    print(f"Wrote fixtures for {', '.join(tickers)} to {FIXTURES}")


if __name__ == "__main__":
    main()
