"""Parse a Fidelity "Portfolio Positions" CSV export into holdings.

Fidelity's export quirks handled here:
  * money as "$1,234.56", "-$12.30", "+$5.00"; percents as "+1.70%"; "--" for not available
  * one row per (account, symbol): the same stock in two accounts is combined
  * money-market sweep rows (symbols ending in "**", e.g. SPAXX**), a "Pending Activity" row,
    option rows (symbols starting with "-"), and legal footer text after the table
  * a UTF-8 byte-order mark at the start of the file

Account numbers are never carried forward; only account *names* are kept.
"""
from __future__ import annotations

import csv
import math
import re
from dataclasses import dataclass, field
from pathlib import Path

CASH, STOCK, OPTION, OTHER = "Cash", "Stock", "Option", "Other"


@dataclass
class Holding:
    symbol: str
    description: str
    kind: str                     # Cash / Stock / Option / Other
    quantity: float = 0.0
    last_price: float = float("nan")
    current_value: float = 0.0
    cost_basis: float = float("nan")      # NaN when Fidelity shows "--" for any lot
    accounts: list[str] = field(default_factory=list)

    @property
    def gain(self) -> float:
        return self.current_value - self.cost_basis

    @property
    def sec_ticker(self) -> str:
        """Fidelity writes class shares as BRK.B; SEC and yfinance use BRK-B."""
        return self.symbol.replace(".", "-")


def parse_money(text: str | None) -> float:
    """'$1,234.56' -> 1234.56, '-$12.30' -> -12.3, '--' or '' -> NaN."""
    if text is None:
        return float("nan")
    cleaned = re.sub(r"[$,+%\s]", "", text)
    if cleaned in ("", "--", "-", "n/a"):
        return float("nan")
    return float(cleaned)


def _add(total: float, x: float) -> float:
    """Sum that treats a missing ("--") value as zero."""
    return total if math.isnan(x) else total + x


def classify(symbol: str, description: str) -> str:
    if symbol.endswith("**") or "MONEY MARKET" in description.upper():
        return CASH
    if symbol.startswith("-"):
        return OPTION
    if re.fullmatch(r"[A-Z]{1,5}([.-][A-Z])?", symbol):
        return STOCK            # may still turn out to be an ETF/fund; the runner finds out from EDGAR
    return OTHER


def _rows(path: Path) -> list[dict[str, str]]:
    """Table rows only: start at the header line, stop at the first blank line (the legal footer follows)."""
    lines = path.read_text(encoding="utf-8-sig").splitlines()
    start = next((i for i, line in enumerate(lines) if ",symbol," in line.lower()), None)
    if start is None:
        raise ValueError(f"{path.name}: no 'Symbol' header row found - is this a Fidelity positions export?")
    table = []
    for line in lines[start:]:
        if not line.strip():
            break
        table.append(line)
    # Fidelity has changed header capitalization over time ("Current Value" vs "Current value"),
    # so columns are matched case-insensitively.
    return [{(k or "").strip().lower(): v for k, v in row.items()} for row in csv.DictReader(table)]


def parse_positions(path: Path | str) -> list[Holding]:
    path = Path(path)
    holdings: dict[str, Holding] = {}
    for row in _rows(path):
        symbol = (row.get("symbol") or "").strip()
        if not symbol or symbol.lower() == "pending activity":
            continue
        description = (row.get("description") or "").strip()
        kind = classify(symbol, description)
        if symbol not in holdings:
            holdings[symbol] = Holding(symbol=symbol.rstrip("*"), description=description, kind=kind, cost_basis=0.0)
        h = holdings[symbol]

        h.current_value = _add(h.current_value, parse_money(row.get("current value")))
        h.quantity = _add(h.quantity, parse_money(row.get("quantity")))
        if h.kind != CASH:
            # Deliberately NaN-propagating: if any lot's cost basis is "--", the total is unknown.
            h.cost_basis += parse_money(row.get("cost basis total"))
        price = parse_money(row.get("last price"))
        if not math.isnan(price):
            h.last_price = price
        account = (row.get("account name") or "").strip()
        if account and account not in h.accounts:
            h.accounts.append(account)

    for h in holdings.values():
        if h.kind == CASH:
            h.cost_basis, h.last_price, h.quantity = h.current_value, 1.0, h.current_value
    return list(holdings.values())
