"""Deprecated backend boundary for the legacy Excel compatibility model."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from copy import deepcopy
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any, Protocol

from xlsliberator.legacy_agent.runtime.errors import ExcelError


class RuntimeCapability(StrEnum):
    CELLS = "excel.cells"
    NAMES = "excel.names"
    TABLES = "excel.tables"
    FILTERS = "excel.filters"
    CHARTS = "excel.charts"
    PIVOTS = "excel.pivots"
    CONTROLS = "excel.controls"
    EVENTS = "excel.events"
    CALCULATION = "excel.calculation"


class CompatibilityBackend(Protocol):
    """Operations required by the typed Excel object model."""

    backend_kind: str

    @property
    def capabilities(self) -> frozenset[RuntimeCapability]: ...

    def workbook_names(self) -> list[str]: ...

    def worksheet_names(self, workbook: str) -> list[str]: ...

    def active_worksheet_name(self, workbook: str) -> str: ...

    def activate_worksheet(self, workbook: str, worksheet: str) -> None: ...

    def read_range(self, workbook: str, worksheet: str, address: str) -> list[list[Any]]: ...

    def write_range(
        self, workbook: str, worksheet: str, address: str, values: Sequence[Sequence[Any]]
    ) -> None: ...

    def read_formula(self, workbook: str, worksheet: str, address: str) -> str | None: ...

    def write_formula(self, workbook: str, worksheet: str, address: str, formula: str) -> None: ...

    def named_items(self, workbook: str) -> dict[str, str]: ...

    def collection_items(
        self, workbook: str, worksheet: str, collection: RuntimeCapability
    ) -> list[dict[str, Any]]: ...

    def recalculate(self, workbook: str) -> None: ...

    def emit_event(self, workbook: str, event_name: str, arguments: Mapping[str, Any]) -> bool: ...


@dataclass
class FakeExcelBackend:
    """Pure deterministic in-memory backend used by conformance tests."""

    backend_kind: str = "fake"
    active_workbook: str = "Book1"
    sheets: dict[str, list[str]] = field(default_factory=lambda: {"Book1": ["Sheet1"]})
    values: dict[tuple[str, str, str], list[list[Any]]] = field(default_factory=dict)
    formulas: dict[tuple[str, str, str], str] = field(default_factory=dict)
    names: dict[str, dict[str, str]] = field(default_factory=dict)
    collections: dict[tuple[str, str, RuntimeCapability], list[dict[str, Any]]] = field(
        default_factory=dict
    )
    events: list[dict[str, Any]] = field(default_factory=list)
    cancelled_events: set[str] = field(default_factory=set)
    calculation_count: int = 0
    active_sheets: dict[str, str] = field(default_factory=dict)

    @property
    def capabilities(self) -> frozenset[RuntimeCapability]:
        return frozenset(RuntimeCapability)

    def workbook_names(self) -> list[str]:
        return list(self.sheets)

    def worksheet_names(self, workbook: str) -> list[str]:
        self._require_workbook(workbook)
        return list(self.sheets[workbook])

    def active_worksheet_name(self, workbook: str) -> str:
        self._require_workbook(workbook)
        return self.active_sheets.get(workbook, self.sheets[workbook][0])

    def activate_worksheet(self, workbook: str, worksheet: str) -> None:
        self._require_sheet(workbook, worksheet)
        self.active_sheets[workbook] = worksheet

    def read_range(self, workbook: str, worksheet: str, address: str) -> list[list[Any]]:
        self._require_sheet(workbook, worksheet)
        return deepcopy(self.values.get((workbook, worksheet, address), [[None]]))

    def write_range(
        self, workbook: str, worksheet: str, address: str, values: Sequence[Sequence[Any]]
    ) -> None:
        self._require_sheet(workbook, worksheet)
        rows = [list(row) for row in values]
        if not rows or any(len(row) != len(rows[0]) for row in rows):
            raise ExcelError("range values must be a non-empty rectangular matrix")
        self.values[(workbook, worksheet, address)] = deepcopy(rows)

    def read_formula(self, workbook: str, worksheet: str, address: str) -> str | None:
        self._require_sheet(workbook, worksheet)
        return self.formulas.get((workbook, worksheet, address))

    def write_formula(self, workbook: str, worksheet: str, address: str, formula: str) -> None:
        self._require_sheet(workbook, worksheet)
        self.formulas[(workbook, worksheet, address)] = formula

    def named_items(self, workbook: str) -> dict[str, str]:
        self._require_workbook(workbook)
        return dict(self.names.get(workbook, {}))

    def collection_items(
        self, workbook: str, worksheet: str, collection: RuntimeCapability
    ) -> list[dict[str, Any]]:
        self._require_sheet(workbook, worksheet)
        if collection not in {
            RuntimeCapability.TABLES,
            RuntimeCapability.FILTERS,
            RuntimeCapability.CHARTS,
            RuntimeCapability.PIVOTS,
            RuntimeCapability.CONTROLS,
        }:
            raise ExcelError(f"{collection.value} is not a worksheet collection")
        return deepcopy(self.collections.get((workbook, worksheet, collection), []))

    def recalculate(self, workbook: str) -> None:
        self._require_workbook(workbook)
        self.calculation_count += 1

    def emit_event(self, workbook: str, event_name: str, arguments: Mapping[str, Any]) -> bool:
        self._require_workbook(workbook)
        self.events.append(
            {"workbook": workbook, "event": event_name, "arguments": dict(arguments)}
        )
        return event_name not in self.cancelled_events

    def _require_workbook(self, workbook: str) -> None:
        if workbook not in self.sheets:
            raise ExcelError(f"workbook not found: {workbook}")

    def _require_sheet(self, workbook: str, worksheet: str) -> None:
        self._require_workbook(workbook)
        if worksheet not in self.sheets[workbook]:
            raise ExcelError(f"worksheet not found: {workbook}.{worksheet}")


class UnoExcelBackend:
    """Real target backend over an injected LibreOffice UNO document.

    This module does not import ``uno``. Construction is allowed only inside the
    guarded office worker, which supplies the already-open target document.
    """

    backend_kind = "libreoffice_uno"

    def __init__(self, document: Any, *, workbook_name: str = "ThisWorkbook") -> None:
        if document is None:
            raise ExcelError("UNO document is required")
        self._document = document
        self._workbook_name = workbook_name
        self._event_log: list[dict[str, Any]] = []

    @property
    def capabilities(self) -> frozenset[RuntimeCapability]:
        return frozenset(
            {
                RuntimeCapability.CELLS,
                RuntimeCapability.NAMES,
                RuntimeCapability.TABLES,
                RuntimeCapability.FILTERS,
                RuntimeCapability.CHARTS,
                RuntimeCapability.PIVOTS,
                RuntimeCapability.CONTROLS,
                RuntimeCapability.EVENTS,
                RuntimeCapability.CALCULATION,
            }
        )

    def workbook_names(self) -> list[str]:
        return [self._workbook_name]

    def worksheet_names(self, workbook: str) -> list[str]:
        self._require_workbook(workbook)
        sheets = self._document.getSheets()
        return [str(name) for name in sheets.getElementNames()]

    def active_worksheet_name(self, workbook: str) -> str:
        self._require_workbook(workbook)
        return str(self._document.getCurrentController().getActiveSheet().getName())

    def activate_worksheet(self, workbook: str, worksheet: str) -> None:
        self._require_workbook(workbook)
        self._document.getCurrentController().setActiveSheet(self._sheet(worksheet))

    def read_range(self, workbook: str, worksheet: str, address: str) -> list[list[Any]]:
        target = self._range(workbook, worksheet, address)
        return [list(row) for row in target.getDataArray()]

    def write_range(
        self, workbook: str, worksheet: str, address: str, values: Sequence[Sequence[Any]]
    ) -> None:
        rows = tuple(tuple(row) for row in values)
        if not rows or any(len(row) != len(rows[0]) for row in rows):
            raise ExcelError("range values must be a non-empty rectangular matrix")
        target = self._range(workbook, worksheet, address)
        if target.getRows().getCount() != len(rows) or target.getColumns().getCount() != len(
            rows[0]
        ):
            raise ExcelError("range values do not match the target range dimensions")
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                cell = target.getCellByPosition(column_index, row_index)
                if value is None:
                    cell.setString("")
                elif isinstance(value, (bool, int, float)):
                    cell.setValue(float(value))
                else:
                    cell.setString(str(value))

    def read_formula(self, workbook: str, worksheet: str, address: str) -> str | None:
        formula = str(self._range(workbook, worksheet, address).getFormula())
        return formula or None

    def write_formula(self, workbook: str, worksheet: str, address: str, formula: str) -> None:
        self._range(workbook, worksheet, address).setFormula(formula)

    def named_items(self, workbook: str) -> dict[str, str]:
        self._require_workbook(workbook)
        names = self._document.getPropertyValue("NamedRanges")
        return {
            str(name): str(names.getByName(name).getContent()) for name in names.getElementNames()
        }

    def collection_items(
        self, workbook: str, worksheet: str, collection: RuntimeCapability
    ) -> list[dict[str, Any]]:
        self._require_workbook(workbook)
        sheet = self._sheet(worksheet)
        if collection is RuntimeCapability.TABLES or collection is RuntimeCapability.FILTERS:
            database_ranges = self._document.getPropertyValue("DatabaseRanges")
            return [{"name": str(name)} for name in database_ranges.getElementNames()]
        if collection is RuntimeCapability.CHARTS:
            charts = sheet.getCharts()
            return [{"name": str(name)} for name in charts.getElementNames()]
        if collection is RuntimeCapability.PIVOTS:
            tables = sheet.getDataPilotTables()
            return [{"name": str(name)} for name in tables.getElementNames()]
        if collection is RuntimeCapability.CONTROLS:
            controls: list[dict[str, Any]] = []
            forms = sheet.getDrawPage().getForms()
            for form_index in range(forms.getCount()):
                form = forms.getByIndex(form_index)
                for control_index in range(form.getCount()):
                    control = form.getByIndex(control_index)
                    controls.append({"name": str(getattr(control, "Name", ""))})
            return controls
        raise ExcelError(f"unsupported UNO collection: {collection.value}")

    def recalculate(self, workbook: str) -> None:
        self._require_workbook(workbook)
        self._document.calculateAll()

    def emit_event(self, workbook: str, event_name: str, arguments: Mapping[str, Any]) -> bool:
        self._require_workbook(workbook)
        self._event_log.append(
            {"workbook": workbook, "event": event_name, "arguments": dict(arguments)}
        )
        return True

    def _require_workbook(self, workbook: str) -> None:
        if workbook != self._workbook_name:
            raise ExcelError(f"workbook not found: {workbook}")

    def _sheet(self, worksheet: str) -> Any:
        sheets = self._document.getSheets()
        if not sheets.hasByName(worksheet):
            raise ExcelError(f"worksheet not found: {worksheet}")
        return sheets.getByName(worksheet)

    def _range(self, workbook: str, worksheet: str, address: str) -> Any:
        self._require_workbook(workbook)
        return self._sheet(worksheet).getCellRangeByName(address)
