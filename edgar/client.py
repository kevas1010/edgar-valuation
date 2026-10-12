"""SEC EDGAR client: polite HTTP fetching with a rate limit and an on-disk JSON cache.

SEC fair-access rules: every request must carry a descriptive User-Agent, and
clients must stay under 10 requests/second. Every response is cached to disk,
so reruns (and the test suite) work fully offline.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import requests

USER_AGENT = "Keval Shah kevalshah505@gmail.com"
TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json"

DEFAULT_CACHE_DIR = Path(__file__).resolve().parent.parent / "cache"
MIN_INTERVAL_S = 0.11  # a little slower than 10 req/s, to leave headroom


class EdgarClient:
    def __init__(self, cache_dir: Path | str = DEFAULT_CACHE_DIR, offline: bool = False):
        self.cache_dir = Path(cache_dir)
        self.offline = offline  # True = never touch the network (used by tests)
        self._session = requests.Session()
        self._session.headers.update({"User-Agent": USER_AGENT, "Accept-Encoding": "gzip, deflate"})
        self._last_request = 0.0

    # ---------- low-level ----------
    def _throttle(self) -> None:
        wait = MIN_INTERVAL_S - (time.monotonic() - self._last_request)
        if wait > 0:
            time.sleep(wait)
        self._last_request = time.monotonic()

    def get_json(self, url: str, cache_name: str) -> dict:
        """Return the JSON at `url`, reading from / writing to cache/<cache_name>."""
        path = self.cache_dir / cache_name
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
        if self.offline:
            raise FileNotFoundError(f"Offline mode and no cached copy of {url} at {path}")

        for attempt in range(3):
            self._throttle()
            resp = self._session.get(url, timeout=30)
            if resp.status_code == 429 or resp.status_code >= 500:
                time.sleep(2 ** attempt)  # back off and retry
                continue
            resp.raise_for_status()
            data = resp.json()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data), encoding="utf-8")
            return data
        resp.raise_for_status()
        raise RuntimeError(f"Failed to fetch {url}")

    # ---------- EDGAR endpoints ----------
    def ticker_to_cik(self, ticker: str) -> int:
        table = self.get_json(TICKERS_URL, "company_tickers.json")
        ticker = ticker.upper()
        for row in table.values():
            if row["ticker"].upper() == ticker:
                return int(row["cik_str"])
        raise ValueError(f"Ticker {ticker!r} not found in SEC company_tickers.json")

    def company_facts(self, ticker: str) -> dict:
        cik = self.ticker_to_cik(ticker)
        return self.get_json(FACTS_URL.format(cik=cik), f"companyfacts/CIK{cik:010d}.json")
