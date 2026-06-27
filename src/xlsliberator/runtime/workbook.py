"""Workbook and worksheet compatibility adapters."""

from __future__ import annotations

from dataclasses import dataclass, field

from xlsliberator.runtime.range import RangeAdapter, row_col_to_address


@dataclass
class WorksheetAdapter:
    """Mockable worksheet adapter."""

    name: str

    def range(self, address: str) -> RangeAdapter:
        """Return a range adapter for an A1 address."""
        return RangeAdapter(self.name, address)

    def cells(self, row: int, col: int) -> RangeAdapter:
        """Return a range adapter for 1-based row/column coordinates."""
        return self.range(row_col_to_address(row, col))


@dataclass
class WorkbookAdapter:
    """Mockable workbook adapter."""

    worksheets: dict[str, WorksheetAdapter] = field(default_factory=dict)
    active_sheet_name: str | None = None

    def worksheet(self, name: str) -> WorksheetAdapter:
        """Return a worksheet adapter, creating a pure adapter if needed."""
        if name not in self.worksheets:
            self.worksheets[name] = WorksheetAdapter(name)
        if self.active_sheet_name is None:
            self.active_sheet_name = name
        return self.worksheets[name]

    def activate(self, name: str) -> WorksheetAdapter:
        """Mark a worksheet active and return it (VBA ``Worksheets(name).Activate``)."""
        sheet = self.worksheet(name)
        self.active_sheet_name = name
        return sheet

    def active_sheet(self) -> WorksheetAdapter:
        """Return the active worksheet (VBA ``ActiveSheet``).

        The first worksheet referenced becomes active by default, mirroring how a
        workbook always has exactly one active sheet. Falls back to creating a
        default sheet only when the workbook is completely empty.
        """
        if self.active_sheet_name is not None:
            return self.worksheet(self.active_sheet_name)
        if self.worksheets:
            return next(iter(self.worksheets.values()))
        return self.worksheet("Sheet1")
