"""Shared fixtures. Everything here is offline: data comes from tests/fixtures/ (real, trimmed
EDGAR + yfinance responses captured on 2026-10-14 by tools/make_fixtures.py)."""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from functools import partial
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from edgar.client import EdgarClient  # noqa: E402
from excel.writer import write_workbook  # noqa: E402
from model.prices import get_prices  # noqa: E402
from valuate import build_model  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"
FIXTURE_DATE = date(2026, 10, 14)
PEERS = ["TER", "COHU", "FORM", "KLIC"]


@pytest.fixture(scope="session")
def client() -> EdgarClient:
    return EdgarClient(cache_dir=FIXTURES, offline=True)


@pytest.fixture(scope="session")
def price_loader():
    return partial(get_prices, cache_dir=FIXTURES / "prices", as_of=FIXTURE_DATE, offline=True)


@pytest.fixture(scope="session")
def aehr_facts(client) -> dict:
    return client.company_facts("AEHR")


@pytest.fixture(scope="session")
def aehr_model(client, price_loader):
    return build_model("AEHR", PEERS, client, price_loader, as_of=FIXTURE_DATE,
                       pulled_at=datetime(2026, 10, 14, 12, 0))


@pytest.fixture(scope="session")
def aehr_workbook(aehr_model, tmp_path_factory) -> Path:
    return write_workbook(aehr_model, tmp_path_factory.mktemp("out") / "AEHR_valuation_test.xlsx")


def load_prices(ticker: str) -> dict:
    return json.loads((FIXTURES / "prices" / f"{ticker}_{FIXTURE_DATE.isoformat()}.json").read_text())
