"""Beta regression: statsmodels OLS must agree with a hand-computed least-squares slope."""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

from model.beta import regress_beta, weekly_returns
from model.prices import weekly_series
from tests.conftest import load_prices


def test_beta_matches_numpy_polyfit_on_real_prices():
    stock, market = weekly_series(load_prices("AEHR")), weekly_series(load_prices("GSPC"))
    result = regress_beta(stock, market)
    slope, intercept = np.polyfit(result.returns["market"], result.returns["stock"], 1)
    assert abs(result.beta - slope) < 1e-6
    assert abs(result.alpha - intercept) < 1e-6
    assert result.n_obs == 104                                   # 2 years of weekly returns
    r = np.corrcoef(result.returns["market"], result.returns["stock"])[0, 1]
    assert result.r_squared == pytest.approx(r ** 2, abs=1e-9)   # simple regression: R² = correlation²


def test_beta_recovers_known_slope_on_synthetic_data():
    rng = np.random.default_rng(42)
    dates = pd.date_range("2024-01-05", periods=105, freq="W-FRI")
    market_ret = rng.normal(0.002, 0.02, len(dates) - 1)
    stock_ret = 0.001 + 1.5 * market_ret + rng.normal(0, 0.01, len(dates) - 1)
    to_prices = lambda r: pd.Series(100 * np.cumprod(np.r_[1.0, 1 + r]), index=dates)

    result = regress_beta(to_prices(stock_ret), to_prices(market_ret))
    slope = np.polyfit(result.returns["market"], result.returns["stock"], 1)[0]
    assert abs(result.beta - slope) < 1e-6
    assert result.beta == pytest.approx(1.5, abs=4 * result.std_err)   # true beta inside ~4 standard errors
    assert result.adjusted_beta == pytest.approx(0.67 * result.beta + 0.33)


def test_returns_use_only_overlapping_dates():
    a = pd.Series([100, 110, 121, 133.1], index=pd.to_datetime(["2024-01-05", "2024-01-12", "2024-01-19", "2024-01-26"]))
    b = pd.Series([50, 55, 60.5], index=pd.to_datetime(["2024-01-12", "2024-01-19", "2024-01-26"]))
    r = weekly_returns(a, b)
    assert len(r) == 2
    assert r["stock"].tolist() == pytest.approx([0.10, 0.10])
