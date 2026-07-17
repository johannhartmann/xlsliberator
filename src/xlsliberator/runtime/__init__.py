"""Compatibility runtime for translated VBA macros."""

from xlsliberator.runtime.backend import (
    CompatibilityBackend,
    FakeExcelBackend,
    RuntimeCapability,
    UnoExcelBackend,
)
from xlsliberator.runtime.context import ExcelContext
from xlsliberator.runtime.errors import ExcelError
from xlsliberator.runtime.object_model import (
    Application,
    Names,
    Range,
    Workbook,
    Workbooks,
    Worksheet,
    WorksheetCollection,
)
from xlsliberator.runtime.worksheet_function import WorksheetFunctionAdapter

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
