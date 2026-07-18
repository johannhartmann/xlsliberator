"""Deprecated Excel-shaped compatibility runtime for the legacy translator."""

from xlsliberator.legacy_agent.runtime.backend import (
    CompatibilityBackend,
    FakeExcelBackend,
    RuntimeCapability,
    UnoExcelBackend,
)
from xlsliberator.legacy_agent.runtime.context import ExcelContext
from xlsliberator.legacy_agent.runtime.errors import ExcelError
from xlsliberator.legacy_agent.runtime.object_model import (
    Application,
    Names,
    Range,
    Workbook,
    Workbooks,
    Worksheet,
    WorksheetCollection,
)
from xlsliberator.legacy_agent.runtime.worksheet_function import WorksheetFunctionAdapter

__all__ = [
    "Application",
    "CompatibilityBackend",
    "ExcelContext",
    "ExcelError",
    "FakeExcelBackend",
    "Names",
    "Range",
    "RuntimeCapability",
    "UnoExcelBackend",
    "Workbook",
    "Workbooks",
    "Worksheet",
    "WorksheetCollection",
    "WorksheetFunctionAdapter",
]
