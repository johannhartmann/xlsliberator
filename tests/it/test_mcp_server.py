"""Integration tests for MCP server.

Tests the FastMCP 2.0 server exposing LibreOffice UNO operations.
"""

import time
from pathlib import Path
from threading import Thread

import pytest

from xlsliberator.mcp_server import mcp


@pytest.fixture
def test_excel_file() -> Path:
    """Path to test Excel file."""
    return Path("tests/data/simple_sheet.xlsx")


@pytest.fixture
def test_ods_file(tmp_path: Path) -> Path:
    """Path to temporary ODS file."""
    return tmp_path / "output.ods"


def test_mcp_server_has_tools() -> None:
    """Test that MCP server has all expected tool functions available."""
    # Verify all expected tool functions can be imported
    from xlsliberator.mcp_tools import (
        compare_formulas,
        convert_excel_to_ods,
        embed_macros,
        get_sheet_data,
        list_sheets,
        read_cell,
        recalculate_document,
        validate_macros,
    )

    # Verify they are callable
    assert callable(convert_excel_to_ods)
    assert callable(recalculate_document)
    assert callable(read_cell)
    assert callable(list_sheets)
    assert callable(get_sheet_data)
    assert callable(compare_formulas)
    assert callable(embed_macros)
    assert callable(validate_macros)


@pytest.mark.integration
@pytest.mark.asyncio
async def test_convert_excel_to_ods_tool(test_excel_file: Path, test_ods_file: Path) -> None:
    """Test convert_excel_to_ods MCP tool."""
    if not test_excel_file.exists():
        pytest.skip(f"Test file not found: {test_excel_file}")

    from xlsliberator.mcp_tools import convert_excel_to_ods

    result = await convert_excel_to_ods(
        excel_path=str(test_excel_file),
        output_path=str(test_ods_file),
        embed_macros=False,  # Skip macros for speed
        use_agent=False,
    )

    assert result["success"], f"Conversion failed: {result.get('error')}"
    assert test_ods_file.exists(), "Output ODS file not created"
    assert "report" in result
    assert result["report"]["sheet_count"] > 0


@pytest.mark.integration
@pytest.mark.asyncio
async def test_list_sheets_tool(test_ods_file: Path) -> None:
    """Test list_sheets MCP tool."""
    if not test_ods_file.exists():
        pytest.skip("ODS file not available")

    from xlsliberator.mcp_tools import list_sheets

    result = await list_sheets(ods_path=str(test_ods_file))

    assert result["success"], f"List sheets failed: {result.get('error')}"
    assert "sheets" in result
    assert isinstance(result["sheets"], list)
    assert result["count"] > 0


@pytest.mark.integration
@pytest.mark.asyncio
async def test_read_cell_tool(test_ods_file: Path) -> None:
    """Test read_cell MCP tool."""
    if not test_ods_file.exists():
        pytest.skip("ODS file not available")

    from xlsliberator.mcp_tools import read_cell

    result = await read_cell(ods_path=str(test_ods_file), sheet_name="Sheet1", cell_address="A1")

    assert result["success"], f"Read cell failed: {result.get('error')}"
    assert "value" in result
    assert "type" in result


def test_mcp_server_metadata() -> None:
    """Test that MCP server has correct metadata."""
    assert mcp.name == "LibreOffice UNO"


@pytest.mark.integration
def test_mcp_server_startup() -> None:
    """Test that MCP server can start and stop gracefully."""
    from xlsliberator.mcp_server import serve

    # Start server in background thread
    server_thread = Thread(target=lambda: serve(host="127.0.0.1", port=8765), daemon=True)
    server_thread.start()

    # Give server time to start
    time.sleep(2)

    # Server should be running
    assert server_thread.is_alive()

    # Note: We can't easily stop the server without modifying serve()
    # but the daemon thread will be cleaned up when the test exits
