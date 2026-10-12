"""Market prices from yfinance, cached to disk as JSON (one file per ticker per day).

Caching per day means reruns on the same day are offline and repeatable,
while a run on a later day picks up fresh prices.
"""
from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pandas as pd

DEFAULT_CACHE_DIR = Path(__file__).resolve().parent.parent / "cache" / "prices"
MARKET_INDEX = "^GSPC"


def _download(ticker: str) -> dict:
    import yfinance as yf  # imported lazily so offline tests never need it

    t = yf.Ticker(ticker)
    weekly = t.history(period="2y", interval="1wk", auto_adjust=True)["Close"].dropna()
    daily = t.history(period="5d", auto_adjust=True)["Close"].dropna()
    if weekly.empty or daily.empty:
        raise ValueError(f"yfinance returned no prices for {ticker!r}")
    return {
        "ticker": ticker,
        "current_price": float(daily.iloc[-1]),
        "price_date": daily.index[-1].strftime("%Y-%m-%d"),
        "weekly_close": {d.strftime("%Y-%m-%d"): float(v) for d, v in weekly.items()},
    }


def get_prices(ticker: str, cache_dir: Path | str = DEFAULT_CACHE_DIR, as_of: date | None = None,
               offline: bool = False) -> dict:
    """{'current_price', 'price_date', 'weekly_close': {date: close}} for `ticker`."""
    as_of = as_of or date.today()
    path = Path(cache_dir) / f"{ticker.replace('^', '')}_{as_of.isoformat()}.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    if offline:
        raise FileNotFoundError(f"Offline mode and no cached prices for {ticker} at {path}")
    data = _download(ticker)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=1), encoding="utf-8")
    return data


def weekly_series(prices: dict) -> pd.Series:
    s = pd.Series(prices["weekly_close"], dtype=float)
    s.index = pd.to_datetime(s.index)
    return s.sort_index()
