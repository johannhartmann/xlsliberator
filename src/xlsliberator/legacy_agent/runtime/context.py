"""Deprecated legacy Excel runtime context."""

from dataclasses import dataclass, field

from xlsliberator.legacy_agent.runtime.backend import CompatibilityBackend, FakeExcelBackend
from xlsliberator.legacy_agent.runtime.object_model import Application, Workbook, Worksheet
from xlsliberator.legacy_agent.runtime.worksheet_function import WorksheetFunctionAdapter


@dataclass
class ExcelContext:
    """Context object passed to translated VBA code."""

    backend: CompatibilityBackend = field(default_factory=FakeExcelBackend)
    application: Application = field(init=False)

    def __post_init__(self) -> None:
        self.application = Application(self.backend)

    @property
    def workbook(self) -> Workbook:
        return self.application.active_workbook

    @property
    def worksheet_function(self) -> WorksheetFunctionAdapter:
        return self.application.worksheet_function

    def active_sheet(self) -> Worksheet:
        """Return the workbook's active sheet (VBA ``ActiveSheet``)."""
        return self.application.active_sheet
