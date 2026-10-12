"""Equity beta from an OLS regression of weekly stock returns on weekly market returns.

    r_stock = alpha + beta * r_market + error

beta is the slope: how many % the stock moves per 1% market move. It feeds
CAPM: cost of equity = risk-free rate + beta * equity risk premium.
"""
from __future__ import annotations

from dataclasses import dataclass

import pandas as pd
import statsmodels.api as sm


@dataclass
class BetaResult:
    beta: float
    alpha: float            # weekly intercept (excess return unexplained by the market)
    r_squared: float        # share of the stock's variance explained by the market
    std_err: float          # standard error of the beta estimate
    n_obs: int
    returns: pd.DataFrame   # columns: stock, market (weekly simple returns) - the chart data

    @property
    def adjusted_beta(self) -> float:
        """Blume adjustment (Bloomberg's 'adjusted beta'): betas drift toward 1 over time."""
        return 0.67 * self.beta + 0.33

    def fit_line(self) -> pd.DataFrame:
        """Two points spanning the observed market returns, for drawing the regression line."""
        x = pd.Series([self.returns["market"].min(), self.returns["market"].max()])
        return pd.DataFrame({"market": x, "fitted": self.alpha + self.beta * x})


def weekly_returns(stock_close: pd.Series, market_close: pd.Series) -> pd.DataFrame:
    """Align both price series on common dates, then compute simple weekly returns."""
    prices = pd.concat({"stock": stock_close, "market": market_close}, axis=1, join="inner")
    return prices.pct_change().dropna()


def regress_beta(stock_close: pd.Series, market_close: pd.Series) -> BetaResult:
    returns = weekly_returns(stock_close, market_close)
    if len(returns) < 10:
        raise ValueError(f"Only {len(returns)} overlapping weekly returns; need at least 10 for a regression.")

    X = sm.add_constant(returns["market"])
    fit = sm.OLS(returns["stock"], X).fit()
    return BetaResult(
        beta=float(fit.params["market"]),
        alpha=float(fit.params["const"]),
        r_squared=float(fit.rsquared),
        std_err=float(fit.bse["market"]),
        n_obs=int(fit.nobs),
        returns=returns,
    )
