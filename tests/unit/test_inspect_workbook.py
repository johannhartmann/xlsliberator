"""Tests for workbook inspection inventory."""

from pathlib import Path
from typing import Any

import openpyxl
from click.testing import CliRunner

from xlsliberator.cli import cli
from xlsliberator.inspect_workbook import inspect_workbook


def _create_formula_workbook(path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = 1
    sheet["A2"] = 2
    sheet["A3"] = "=SUM(A1:A2)"
    workbook.save(path)
    workbook.close()


def test_inspect_workbook_collects_xlsx_formulas(tmp_path: Path) -> None:
    """Inspection should inventory formulas from extracted workbook IR."""
    workbook_path = tmp_path / "sample.xlsx"
    _create_formula_workbook(workbook_path)

    inventory = inspect_workbook(workbook_path)

    assert inventory.workbook.sheet_count == 1
    assert len(inventory.formulas) == 1
    assert inventory.formulas[0].formula_text == "=SUM(A1:A2)"
    assert inventory.unsupported_artifacts == []


def test_inspect_workbook_reports_xls_incomplete(tmp_path: Path, monkeypatch: Any) -> None:
    """Legacy XLS should report incomplete BIFF parsing explicitly."""
    workbook_path = tmp_path / "legacy.xls"
    workbook_path.write_bytes(b"not a real xls")

    from xlsliberator import inspect_workbook as inspect_module
    from xlsliberator.ir_models import ExtractionStats, WorkbookIR

    monkeypatch.setattr(
        inspect_module,
        "extract_workbook",
        lambda path: (WorkbookIR(file_path=str(path), file_format="xls"), ExtractionStats()),
    )
    monkeypatch.setattr(inspect_module, "extract_vba_modules", lambda _path: [])

    inventory = inspect_module.inspect_workbook(workbook_path)

    reasons = [artifact.reason for artifact in inventory.unsupported_artifacts]
    assert any("legacy XLS BIFF parsing incomplete" in reason for reason in reasons)


def test_inspect_cli_json(tmp_path: Path) -> None:
    """CLI inspection should print structured JSON."""
    workbook_path = tmp_path / "sample.xlsx"
    _create_formula_workbook(workbook_path)

    result = CliRunner().invoke(cli, ["inspect", str(workbook_path), "--json"])

    assert result.exit_code == 0
    assert '"formulas"' in result.output
    assert "=SUM(A1:A2)" in result.output
