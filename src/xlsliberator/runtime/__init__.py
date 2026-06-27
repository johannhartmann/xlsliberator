"""Compatibility runtime for translated VBA macros."""

from xlsliberator.runtime.context import ExcelContext
from xlsliberator.runtime.errors import ExcelError
from xlsliberator.runtime.range import RangeAdapter
from xlsliberator.runtime.workbook import WorkbookAdapter, WorksheetAdapter
from xlsliberator.runtime.worksheet_function import WorksheetFunctionAdapter

__all__ = [
    "ExcelContext",
    "ExcelError",
    "RangeAdapter",
    "WorkbookAdapter",
    "WorksheetAdapter",
    "WorksheetFunctionAdapter",
]
