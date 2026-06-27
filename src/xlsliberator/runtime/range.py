"""Range adapter primitives for translated VBA."""

from __future__ import annotations

import re
from dataclasses import dataclass

from xlsliberator.runtime.errors import ExcelError


@dataclass(frozen=True)
class RangeAdapter:
    """Pure address adapter for a spreadsheet range."""

    worksheet_name: str
    address: str
    value: object | None = None
    formula: str | None = None

    def offset(self, row_delta: int = 0, col_delta: int = 0) -> RangeAdapter:
        """Return a new range offset from the top-left cell."""
        row, col = address_to_row_col(self.address)
        return RangeAdapter(
            worksheet_name=self.worksheet_name,
            address=row_col_to_address(row + row_delta, col + col_delta),
            value=self.value,
            formula=self.formula,
        )


def normalize_address(address: str) -> str:
    """Normalize a single-cell A1 address."""
    row, col = address_to_row_col(address)
    return row_col_to_address(row, col)


def address_to_row_col(address: str) -> tuple[int, int]:
    """Convert A1 address to 1-based row/column."""
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?([0-9]+)", address)
    if not match:
        raise ExcelError(f"Unsupported address: {address}")
    col_letters, row_text = match.groups()
    col = 0
    for char in col_letters.upper():
        col = col * 26 + (ord(char) - ord("A") + 1)
    return int(row_text), col


def row_col_to_address(row: int, col: int) -> str:
    """Convert 1-based row/column to A1 address."""
    if row < 1 or col < 1:
        raise ExcelError(f"Invalid row/column: {row}, {col}")
    letters = ""
    value = col
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row}"
