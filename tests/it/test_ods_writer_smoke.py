"""Integration tests for ODS writer (Phase F4 - Gate G4)."""

import tempfile
from pathlib import Path

import pytest

from xlsliberator.ir_models import CellIR, CellType, SheetIR, WorkbookIR
from xlsliberator.uno_conn import UnoCtx, get_cell, get_sheet, open_calc, recalc
from xlsliberator.write_ods import build_calc_from_ir, write_ods_from_ir


@pytest.fixture
def skip_if_no_lo(skip_if_no_lo: None) -> None:
    """Skip tests if LibreOffice is not available."""
    pass


def create_test_ir_with_formulas() -> WorkbookIR:
    """Create test WorkbookIR with 10 formula types (Gate G4 requirement).

    Returns:
        WorkbookIR with test data and formulas
    """
    sheet = SheetIR(name="TestSheet", index=0)

    # Add input values
    sheet.cells.append(CellIR(row=0, col=0, address="A1", cell_type=CellType.NUMBER, value=10))
    sheet.cells.append(CellIR(row=1, col=0, address="A2", cell_type=CellType.NUMBER, value=20))
    sheet.cells.append(CellIR(row=2, col=0, address="A3", cell_type=CellType.NUMBER, value=30))
    sheet.cells.append(CellIR(row=0, col=1, address="B1", cell_type=CellType.STRING, value="Hello"))
    sheet.cells.append(CellIR(row=1, col=1, address="B2", cell_type=CellType.STRING, value="World"))

    # Add 10 formulas (one for each supported function)
    # 1. SUM
    sheet.cells.append(
        CellIR(row=4, col=0, address="A5", cell_type=CellType.FORMULA, formula="=SUM(A1:A3)")
    )

    # 2. AVERAGE
    sheet.cells.append(
        CellIR(row=5, col=0, address="A6", cell_type=CellType.FORMULA, formula="=AVERAGE(A1:A3)")
    )

    # 3. Simple addition
    sheet.cells.append(
        CellIR(row=6, col=0, address="A7", cell_type=CellType.FORMULA, formula="=A1+A2")
    )

    # 4. Multiplication
    sheet.cells.append(
        CellIR(row=7, col=0, address="A8", cell_type=CellType.FORMULA, formula="=A1*2")
    )

    # 5. Division
    sheet.cells.append(
        CellIR(row=8, col=0, address="A9", cell_type=CellType.FORMULA, formula="=A2/2")
    )

    # 6. Subtraction
    sheet.cells.append(
        CellIR(row=9, col=0, address="A10", cell_type=CellType.FORMULA, formula="=A3-A1")
    )

    # 7. Mixed operations
    sheet.cells.append(
        CellIR(row=10, col=0, address="A11", cell_type=CellType.FORMULA, formula="=A1+A2*2")
    )

    # 8. Parentheses
    sheet.cells.append(
        CellIR(row=11, col=0, address="A12", cell_type=CellType.FORMULA, formula="=(A1+A2)*2")
    )

    # 9. Nested SUM (using semicolon for LibreOffice)
    sheet.cells.append(
        CellIR(row=12, col=0, address="A13", cell_type=CellType.FORMULA, formula="=SUM(A1;A2)+A3")
    )

    # 10. AVERAGE with multiplication
    sheet.cells.append(
        CellIR(
            row=13, col=0, address="A14", cell_type=CellType.FORMULA, formula="=AVERAGE(A1:A2)*2"
        )
    )

    wb_ir = WorkbookIR(
        file_path="test.xlsx",
        file_format="xlsx",
        sheets=[sheet],
    )

    return wb_ir


@pytest.mark.integration
def test_build_calc_from_ir(skip_if_no_lo: None) -> None:
    """Test building Calc document from IR."""
    wb_ir = create_test_ir_with_formulas()

    with UnoCtx() as ctx:
        doc = build_calc_from_ir(ctx, wb_ir, locale="en-US")

        # Verify document was created
        assert doc is not None

        # Verify sheets
        sheets = doc.getSheets()
        assert sheets.getCount() >= 1

        sheet = get_sheet(ctx, doc, "TestSheet")
        assert sheet is not None

        # Close document
        doc.close(True)


@pytest.mark.integration
def test_formulas_recalc_correctly(skip_if_no_lo: None) -> None:
    """Test that formulas recalculate correctly (Gate G4 requirement).

    Verifies 10/10 formulas produce correct values within ±1e-9 tolerance.
    """
    wb_ir = create_test_ir_with_formulas()

    with UnoCtx() as ctx:
        doc = build_calc_from_ir(ctx, wb_ir, locale="en-US")
        sheet = get_sheet(ctx, doc, "TestSheet")

        # Recalculate
        recalc(ctx, doc)

        # Define expected results (Gate G4 requirement)
        expected = {
            "A5": 60.0,  # =SUM(A1:A3) = 10+20+30
            "A6": 20.0,  # =AVERAGE(A1:A3) = 60/3
            "A7": 30.0,  # =A1+A2 = 10+20
            "A8": 20.0,  # =A1*2 = 10*2
            "A9": 10.0,  # =A2/2 = 20/2
            "A10": 20.0,  # =A3-A1 = 30-10
            "A11": 50.0,  # =A1+A2*2 = 10+40
            "A12": 60.0,  # =(A1+A2)*2 = 30*2
            "A13": 60.0,  # =SUM(A1,A2)+A3 = 30+30
            "A14": 30.0,  # =AVERAGE(A1:A2)*2 = 15*2
        }

        # Verify each formula result (Gate G4 tolerance: ±1e-9)
        tolerance = 1e-9
        passed = 0
        failed = []

        for address, expected_value in expected.items():
            cell = get_cell(ctx, sheet, address)

            # Check if it's a formula
            formula = cell.getFormula()
            assert formula.startswith("="), f"Cell {address} should have a formula"

            # Get calculated value
            actual_value = cell.getValue()
            diff = abs(actual_value - expected_value)

            if diff <= tolerance:
                passed += 1
            else:
                failed.append(
                    f"{address}: expected {expected_value}, got {actual_value}, diff={diff}"
                )

        # Gate G4 requirement: 10/10 formulas correct
        assert passed == 10, f"Expected 10/10 formulas correct, got {passed}/10. Failed: {failed}"

        doc.close(True)


@pytest.mark.integration
def test_write_and_reopen_ods(skip_if_no_lo: None) -> None:
    """Test writing ODS file and reopening it."""
    wb_ir = create_test_ir_with_formulas()

    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.ods"

        with UnoCtx() as ctx:
            # Write ODS
            write_ods_from_ir(ctx, wb_ir, str(output_path), locale="en-US")

            # Verify file was created
            assert output_path.exists()
            assert output_path.stat().st_size > 0

            # Reopen and verify
            doc = open_calc(ctx, output_path)
            sheet = get_sheet(ctx, doc, "TestSheet")

            # Check a few key values
            cell_a1 = get_cell(ctx, sheet, "A1")
            assert cell_a1.getValue() == 10.0

            cell_a5 = get_cell(ctx, sheet, "A5")
            assert cell_a5.getFormula() == "=SUM(A1:A3)"

            # Recalc and check result
            recalc(ctx, doc)
            assert abs(cell_a5.getValue() - 60.0) < 1e-9

            doc.close(True)


@pytest.mark.integration
def test_multiple_sheets(skip_if_no_lo: None) -> None:
    """Test creating document with multiple sheets."""
    # Create IR with 2 sheets
    sheet1 = SheetIR(name="Sheet1", index=0)
    sheet1.cells.append(CellIR(row=0, col=0, address="A1", cell_type=CellType.NUMBER, value=100))

    sheet2 = SheetIR(name="Sheet2", index=1)
    sheet2.cells.append(CellIR(row=0, col=0, address="A1", cell_type=CellType.NUMBER, value=200))

    wb_ir = WorkbookIR(
        file_path="test.xlsx",
        file_format="xlsx",
        sheets=[sheet1, sheet2],
    )

    with UnoCtx() as ctx:
        doc = build_calc_from_ir(ctx, wb_ir, locale="en-US")

        # Verify both sheets exist
        sheets = doc.getSheets()
        assert sheets.getCount() >= 2

        # Verify sheet names
        s1 = get_sheet(ctx, doc, "Sheet1")
        assert s1 is not None

        s2 = get_sheet(ctx, doc, "Sheet2")
        assert s2 is not None

        # Verify values
        cell_s1_a1 = get_cell(ctx, s1, "A1")
        assert cell_s1_a1.getValue() == 100.0

        cell_s2_a1 = get_cell(ctx, s2, "A1")
        assert cell_s2_a1.getValue() == 200.0

        doc.close(True)


@pytest.mark.integration
def test_different_cell_types(skip_if_no_lo: None) -> None:
    """Test writing different cell types (numbers, strings, booleans)."""
    sheet = SheetIR(name="TestSheet", index=0)

    # Add different cell types
    sheet.cells.append(CellIR(row=0, col=0, address="A1", cell_type=CellType.NUMBER, value=42.5))
    sheet.cells.append(CellIR(row=1, col=0, address="A2", cell_type=CellType.STRING, value="Test"))
    sheet.cells.append(CellIR(row=2, col=0, address="A3", cell_type=CellType.BOOLEAN, value=True))

    wb_ir = WorkbookIR(
        file_path="test.xlsx",
        file_format="xlsx",
        sheets=[sheet],
    )

    with UnoCtx() as ctx:
        doc = build_calc_from_ir(ctx, wb_ir, locale="en-US")
        sheet_uno = get_sheet(ctx, doc, "TestSheet")

        # Verify number
        cell_a1 = get_cell(ctx, sheet_uno, "A1")
        assert cell_a1.getValue() == 42.5

        # Verify string
        cell_a2 = get_cell(ctx, sheet_uno, "A2")
        assert cell_a2.getString() == "Test"

        # Verify boolean (stored as number)
        cell_a3 = get_cell(ctx, sheet_uno, "A3")
        assert cell_a3.getValue() == 1.0

        doc.close(True)
