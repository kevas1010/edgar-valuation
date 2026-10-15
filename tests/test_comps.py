"""Comparable companies: EV math, n/m exclusions (negative metrics and outliers), implied value."""
from __future__ import annotations

import math

import pytest

from edgar.extract import DataNote
from model.comps import MAX_MULTIPLE, PeerData, implied_value_per_share, load_peers, multiples_table


def peer(ticker, ebitda=100.0, net_income=50.0, price=10.0, shares=100.0, debt=200.0, cash=100.0, revenue=500.0):
    return PeerData(ticker, ticker, 2025, price, shares, debt, cash, revenue, ebitda, net_income)


def test_market_cap_ev_and_multiples():
    p = peer("A")                              # mcap 1000, EV = 1000 + 200 − 100 = 1100
    assert p.market_cap == 1000 and p.enterprise_value == 1100
    assert p.multiple("EV/Revenue") == pytest.approx(2.2)
    assert p.multiple("EV/EBITDA") == pytest.approx(11.0)
    assert p.multiple("P/E") == pytest.approx(20.0)


def test_negative_and_outlier_multiples_are_excluded():
    loss = peer("LOSS", ebitda=-5.0, net_income=-1.0)
    trough = peer("TROUGH", net_income=0.5)          # P/E = 2000x > cap
    assert math.isnan(loss.multiple("EV/EBITDA")) and math.isnan(loss.multiple("P/E"))
    assert trough.raw_multiple("P/E") > MAX_MULTIPLE and math.isnan(trough.multiple("P/E"))

    table = multiples_table([peer("A"), peer("B", ebitda=50.0), loss, trough])
    assert table.loc["Median", "EV/EBITDA"] == pytest.approx(11.0)   # median of 11x, 22x, 11x (LOSS excluded)
    assert table.loc["Median", "P/E"] == pytest.approx(20.0)         # only A and B; TROUGH is n/m


def test_implied_value_per_share():
    peers = [peer("A"), peer("B", ebitda=50.0)]                       # EV/EBITDA 11x and 22x -> median 16.5x
    v = implied_value_per_share(peers, revenue=1000, ebitda=200, net_income=-10, debt=300, cash=100, shares=50)
    assert v["EV/EBITDA"] == pytest.approx((200 * 16.5 - 300 + 100) / 50)
    assert math.isnan(v["P/E"])                                       # target loss -> n/m
    assert v["Average"] == pytest.approx((v["EV/Revenue"] + v["EV/EBITDA"]) / 2)


def test_bad_peer_ticker_is_skipped_with_note(client, price_loader):
    notes: list[DataNote] = []
    peers = load_peers(client, ["FORM", "NOTAREALTICKER"], notes, price_loader)
    assert [p.ticker for p in peers] == ["FORM"]
    assert any(n.item == "NOTAREALTICKER" and "Skipped" in n.message for n in notes)
