"""Deprecated Excel compatibility object model used only by the legacy translator."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from typing import Any

from xlsliberator.legacy_agent.runtime.backend import CompatibilityBackend, RuntimeCapability
from xlsliberator.legacy_agent.runtime.errors import ExcelError
from xlsliberator.legacy_agent.runtime.range import address_to_row_col, row_col_to_address
from xlsliberator.legacy_agent.runtime.worksheet_function import WorksheetFunctionAdapter


@dataclass
class Range:
    backend: CompatibilityBackend
    workbook_name: str
    worksheet_name: str
    address: str

    @property
    def value(self) -> Any:
        values = self.backend.read_range(self.workbook_name, self.worksheet_name, self.address)
        return values[0][0] if len(values) == 1 and len(values[0]) == 1 else values

    @value.setter
    def value(self, value: Any) -> None:
        matrix = value if _is_matrix(value) else [[value]]
        self.backend.write_range(self.workbook_name, self.worksheet_name, self.address, matrix)

    @property
    def formula(self) -> str | None:
        return self.backend.read_formula(self.workbook_name, self.worksheet_name, self.address)

    @formula.setter
    def formula(self, value: str) -> None:
        self.backend.write_formula(self.workbook_name, self.worksheet_name, self.address, value)

    @property
    def default(self) -> Any:
        """Excel Range default member is Value."""
        return self.value

    def offset(self, row_delta: int = 0, column_delta: int = 0) -> Range:
        row, column = address_to_row_col(self.address)
        return Range(
            self.backend,
            self.workbook_name,
            self.worksheet_name,
            row_col_to_address(row + row_delta, column + column_delta),
        )


@dataclass(frozen=True)
class WorksheetCollection:
    backend: CompatibilityBackend
    workbook_name: str

    def item(self, name_or_index: str | int) -> Worksheet:
        names = self.backend.worksheet_names(self.workbook_name)
        if isinstance(name_or_index, int):
            if name_or_index < 1 or name_or_index > len(names):
                raise ExcelError(f"worksheet index out of range: {name_or_index}")
            name = names[name_or_index - 1]
        else:
            name = name_or_index
            if name not in names:
                raise ExcelError(f"worksheet not found: {name}")
        return Worksheet(self.backend, self.workbook_name, name)

    def __iter__(self):  # type: ignore[no-untyped-def]
        for name in self.backend.worksheet_names(self.workbook_name):
            yield Worksheet(self.backend, self.workbook_name, name)


@dataclass(frozen=True)
class NamedItem:
    name: str
    refers_to: str


@dataclass(frozen=True)
class Names:
    backend: CompatibilityBackend
    workbook_name: str

    def item(self, name: str) -> NamedItem:
        items = self.backend.named_items(self.workbook_name)
        if name not in items:
            raise ExcelError(f"name not found: {name}")
        return NamedItem(name, items[name])


@dataclass(frozen=True)
class WorksheetArtifacts:
    backend: CompatibilityBackend
    workbook_name: str
    worksheet_name: str
    capability: RuntimeCapability

    def items(self) -> list[dict[str, Any]]:
        if self.capability not in self.backend.capabilities:
            raise ExcelError(f"runtime capability unavailable: {self.capability.value}")
        return self.backend.collection_items(
            self.workbook_name, self.worksheet_name, self.capability
        )


@dataclass(frozen=True)
class Worksheet:
    backend: CompatibilityBackend
    workbook_name: str
    name: str

    def range(self, address: str) -> Range:
        return Range(self.backend, self.workbook_name, self.name, address)

    def cells(self, row: int, column: int) -> Range:
        return self.range(_row_column_to_address(row, column))

    def activate(self) -> None:
        self.backend.activate_worksheet(self.workbook_name, self.name)

    @property
    def tables(self) -> WorksheetArtifacts:
        return WorksheetArtifacts(
            self.backend, self.workbook_name, self.name, RuntimeCapability.TABLES
        )

    @property
    def filters(self) -> WorksheetArtifacts:
        return WorksheetArtifacts(
            self.backend, self.workbook_name, self.name, RuntimeCapability.FILTERS
        )

    @property
    def charts(self) -> WorksheetArtifacts:
        return WorksheetArtifacts(
            self.backend, self.workbook_name, self.name, RuntimeCapability.CHARTS
        )

    @property
    def pivots(self) -> WorksheetArtifacts:
        return WorksheetArtifacts(
            self.backend, self.workbook_name, self.name, RuntimeCapability.PIVOTS
        )

    @property
    def controls(self) -> WorksheetArtifacts:
        return WorksheetArtifacts(
            self.backend, self.workbook_name, self.name, RuntimeCapability.CONTROLS
        )


@dataclass(frozen=True)
class Workbook:
    backend: CompatibilityBackend
    name: str

    @property
    def worksheets(self) -> WorksheetCollection:
        return WorksheetCollection(self.backend, self.name)

    @property
    def names(self) -> Names:
        return Names(self.backend, self.name)

    def recalculate(self) -> None:
        self.backend.recalculate(self.name)

    def emit_event(self, event_name: str, arguments: Mapping[str, Any] | None = None) -> bool:
        return self.backend.emit_event(self.name, event_name, arguments or {})


@dataclass(frozen=True)
class Workbooks:
    backend: CompatibilityBackend

    def item(self, name_or_index: str | int) -> Workbook:
        names = self.backend.workbook_names()
        if isinstance(name_or_index, int):
            if name_or_index < 1 or name_or_index > len(names):
                raise ExcelError(f"workbook index out of range: {name_or_index}")
            name = names[name_or_index - 1]
        else:
            name = name_or_index
            if name not in names:
                raise ExcelError(f"workbook not found: {name}")
        return Workbook(self.backend, name)


@dataclass(frozen=True)
class Application:
    backend: CompatibilityBackend
    worksheet_function: WorksheetFunctionAdapter = WorksheetFunctionAdapter()

    @property
    def workbooks(self) -> Workbooks:
        return Workbooks(self.backend)

    @property
    def active_workbook(self) -> Workbook:
        names = self.backend.workbook_names()
        if not names:
            raise ExcelError("no active workbook")
        return Workbook(self.backend, names[0])

    @property
    def active_sheet(self) -> Worksheet:
        workbook = self.active_workbook
        return workbook.worksheets.item(self.backend.active_worksheet_name(workbook.name))

    def calculate(self) -> None:
        self.active_workbook.recalculate()


def _is_matrix(value: Any) -> bool:
    return (
        isinstance(value, Sequence)
        and not isinstance(value, (str, bytes))
        and bool(value)
        and all(isinstance(row, Sequence) and not isinstance(row, (str, bytes)) for row in value)
    )


def _row_column_to_address(row: int, column: int) -> str:
    if row < 1 or column < 1:
        raise ExcelError(f"invalid row/column: {row}, {column}")
    letters = ""
    remaining = column
    while remaining:
        remaining, remainder = divmod(remaining - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row}"
