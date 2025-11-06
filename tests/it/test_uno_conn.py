"""Integration tests for LibreOffice UNO connection."""

import tempfile
from pathlib import Path

import pytest

from xlsliberator.uno_conn import (
    UnoConnectionError,
    UnoCtx,
    connect_lo,
    get_cell,
    get_sheet,
    new_calc,
    recalc,
    save_as_ods,
)


@pytest.fixture
def skip_if_no_lo(skip_if_no_lo: None) -> None:
    """Skip tests if LibreOffice is not available."""
    pass


@pytest.mark.integration
def test_connect_disconnect(skip_if_no_lo: None) -> None:
    """Test basic connection and disconnection."""
    ctx = connect_lo()
    assert ctx.is_connected

    ctx.disconnect()
    assert not ctx.is_connected


@pytest.mark.integration
def test_context_manager(skip_if_no_lo: None) -> None:
    """Test UnoCtx as context manager."""
    with UnoCtx() as ctx:
        assert ctx.is_connected

    # After exiting context, should be disconnected
    assert not ctx.is_connected


@pytest.mark.integration
def test_new_calc_document(skip_if_no_lo: None) -> None:
    """Test creating a new Calc document."""
    with UnoCtx() as ctx:
        doc = new_calc(ctx)
        assert doc is not None

        # Check document has at least one sheet
        sheets = doc.getSheets()
        assert sheets.getCount() >= 1

        # Close document
        doc.close(True)


@pytest.mark.integration
def test_recalc_empty_document(skip_if_no_lo: None) -> None:
    """Test recalculating an empty document (Gate G2 requirement)."""
    with UnoCtx() as ctx:
        doc = new_calc(ctx)

        # Should not raise any errors
        recalc(ctx, doc)

        doc.close(True)


@pytest.mark.integration
def test_get_sheet(skip_if_no_lo: None) -> None:
    """Test getting sheets by name and index."""
    with UnoCtx() as ctx:
        doc = new_calc(ctx)

        # Get sheet by index
        sheet = get_sheet(ctx, doc, 0)
        assert sheet is not None

        # Get sheet by name (default sheet usually named "Sheet1" or similar)
        sheet_name = sheet.getName()
        sheet2 = get_sheet(ctx, doc, sheet_name)
        assert sheet2 is not None

        # Test invalid sheet name
        with pytest.raises(KeyError):
            get_sheet(ctx, doc, "NonExistentSheet")

        # Test invalid sheet index
        with pytest.raises(IndexError):
            get_sheet(ctx, doc, 999)

        doc.close(True)


@pytest.mark.integration
def test_get_cell(skip_if_no_lo: None) -> None:
    """Test getting cells by address."""
    with UnoCtx() as ctx:
        doc = new_calc(ctx)
        sheet = get_sheet(ctx, doc, 0)

        # Get cell A1
        cell = get_cell(ctx, sheet, "A1")
        assert cell is not None

        # Set and read a value
        cell.setString("Test")
        assert cell.getString() == "Test"

        # Get another cell
        cell_b2 = get_cell(ctx, sheet, "B2")
        cell_b2.setValue(42.0)
        assert cell_b2.getValue() == 42.0

        doc.close(True)


@pytest.mark.integration
def test_save_as_ods(skip_if_no_lo: None) -> None:
    """Test saving document as ODS."""
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "test.ods"

        with UnoCtx() as ctx:
            doc = new_calc(ctx)

            # Add some content
            sheet = get_sheet(ctx, doc, 0)
            cell = get_cell(ctx, sheet, "A1")
            cell.setString("Test Content")

            # Save document
            save_as_ods(ctx, doc, output_path)

            doc.close(True)

        # Verify file was created
        assert output_path.exists()
        assert output_path.stat().st_size > 0


@pytest.mark.integration
def test_formula_recalc(skip_if_no_lo: None) -> None:
    """Test formula recalculation."""
    with UnoCtx() as ctx:
        doc = new_calc(ctx)
        sheet = get_sheet(ctx, doc, 0)

        # Set values
        cell_a1 = get_cell(ctx, sheet, "A1")
        cell_a1.setValue(10.0)

        cell_a2 = get_cell(ctx, sheet, "A2")
        cell_a2.setValue(20.0)

        # Set formula
        cell_a3 = get_cell(ctx, sheet, "A3")
        cell_a3.setFormula("=A1+A2")

        # Recalculate
        recalc(ctx, doc)

        # Check result
        assert cell_a3.getValue() == 30.0

        doc.close(True)


@pytest.mark.integration
@pytest.mark.slow
def test_stability_10_cycles(skip_if_no_lo: None) -> None:
    """Test connection stability over 10 connect/disconnect cycles (Gate G2)."""
    for i in range(10):
        with UnoCtx() as ctx:
            assert ctx.is_connected

            # Create a document in each cycle
            doc = new_calc(ctx)

            # Do a simple operation
            sheet = get_sheet(ctx, doc, 0)
            cell = get_cell(ctx, sheet, "A1")
            cell.setValue(float(i))

            # Recalculate
            recalc(ctx, doc)

            # Close document
            doc.close(True)

        # Verify disconnected after context exit
        assert not ctx.is_connected


@pytest.mark.integration
def test_error_when_not_connected() -> None:
    """Test that operations fail when not connected."""
    ctx = UnoCtx()
    assert not ctx.is_connected

    # Should raise error
    with pytest.raises(UnoConnectionError):
        new_calc(ctx)
