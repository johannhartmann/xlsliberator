"""Excel runtime context."""

from dataclasses import dataclass, field

from xlsliberator.runtime.workbook import WorkbookAdapter, WorksheetAdapter
from xlsliberator.runtime.worksheet_function import WorksheetFunctionAdapter


@dataclass
class ExcelContext:
    """Context object passed to translated VBA code."""

    workbook: WorkbookAdapter = field(default_factory=WorkbookAdapter)
    worksheet_function: WorksheetFunctionAdapter = field(default_factory=WorksheetFunctionAdapter)

    def active_sheet(self) -> WorksheetAdapter:
        """Return the workbook's active sheet (VBA ``ActiveSheet``)."""
        return self.workbook.active_sheet()
