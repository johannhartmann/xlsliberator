"""Tests for VBA compatibility runtime skeleton."""

import pytest

from xlsliberator.legacy_agent.runtime.backend import FakeExcelBackend
from xlsliberator.legacy_agent.runtime.errors import ExcelError
from xlsliberator.legacy_agent.runtime.object_model import Application
from xlsliberator.legacy_agent.runtime.range import (
    address_to_row_col,
    normalize_address,
    row_col_to_address,
)


def test_address_normalization() -> None:
    """A1 addresses should normalize predictably."""
    assert normalize_address("$a$1") == "A1"
    assert address_to_row_col("AA10") == (10, 27)
    assert row_col_to_address(10, 27) == "AA10"


def test_range_offset() -> None:
    """Range offsets should calculate target addresses without UNO."""
    application = Application(FakeExcelBackend())
    assert application.active_sheet.range("B2").offset(2, 3).address == "E4"


def test_workbook_worksheet_cells() -> None:
    """Workbook and worksheet adapters should import and create ranges."""
    workbook = Application(FakeExcelBackend()).active_workbook
    sheet = workbook.worksheets.item("Sheet1")

    assert sheet.cells(3, 2).address == "B3"


def test_invalid_offset_raises() -> None:
    """Invalid addresses should fail clearly."""
    with pytest.raises(ExcelError):
        Application(FakeExcelBackend()).active_sheet.range("A1").offset(-1, 0)


def test_active_sheet_follows_workbook_not_hardcoded_name() -> None:
    """ExcelContext.active_sheet should track the workbook, not a hardcoded Sheet1."""
    from xlsliberator.legacy_agent.runtime.context import ExcelContext

    backend = FakeExcelBackend(sheets={"Book1": ["Dashboard", "Summary"]})
    ctx = ExcelContext(backend)
    assert ctx.active_sheet().name == "Dashboard"

    ctx.workbook.worksheets.item("Summary").activate()
    assert ctx.active_sheet().name == "Summary"
