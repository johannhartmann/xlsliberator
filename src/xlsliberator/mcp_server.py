"""FastMCP 2.0 server for LibreOffice UNO operations.

Exposes xlsliberator functionality through Model Context Protocol (MCP)
for integration with Claude Agent SDK and other MCP clients.

Usage:
    # Run with trusted-local HTTP streaming transport
    docker compose run --rm test python -m xlsliberator.mcp_server

    # Or via CLI
    xlsliberator mcp-serve --port 8000

    # Connect client to: http://localhost:8000/mcp
"""

import os
from pathlib import Path

from fastmcp import FastMCP
from loguru import logger

from xlsliberator.container_boundary import require_application_container
from xlsliberator.mcp_tools import (
    click_form_button,
    compare_formulas,
    convert_excel_to_ods,
    embed_macros,
    execute_button_handler,
    get_cell_colors,
    get_sheet_data,
    inspect_workbook,
    list_controls,
    list_embedded_macros,
    list_event_bindings,
    list_sheets,
    open_document_gui,
    read_cell,
    recalculate_document,
    send_keyboard_input,
    take_screenshot,
    test_macro_execution,
    validate_document_runtime,
    validate_macros,
    validate_transformation,
)

# Create FastMCP server instance
mcp = FastMCP(name="LibreOffice UNO")


# ==============================================================================
# Register Tools
# ==============================================================================

# Document Operations
mcp.tool(convert_excel_to_ods)
mcp.tool(inspect_workbook)
mcp.tool(validate_transformation)
mcp.tool(validate_document_runtime)
mcp.tool(recalculate_document)

# Cell and Sheet Operations
mcp.tool(read_cell)
mcp.tool(list_sheets)
mcp.tool(get_sheet_data)
mcp.tool(list_controls)
mcp.tool(list_event_bindings)

# Formula Testing
mcp.tool(compare_formulas)

# Macro Operations
mcp.tool(embed_macros)
mcp.tool(validate_macros)
mcp.tool(list_embedded_macros)
mcp.tool(test_macro_execution)

# GUI Testing Operations
mcp.tool(open_document_gui)
mcp.tool(click_form_button)
mcp.tool(execute_button_handler)
mcp.tool(send_keyboard_input)
mcp.tool(get_cell_colors)
mcp.tool(take_screenshot)


def serve(host: str = "127.0.0.1", port: int = 8000, *, trusted_local: bool = True) -> None:
    """Start the MCP server in explicit trusted-local mode.

    Network exposure is rejected until an authenticated and authorized
    transport is configured. Loopback is the only trusted-local binding.

    Args:
        host: Loopback address to bind to (default: 127.0.0.1)
        port: Port number (default: 8000)
    """
    require_application_container()
    if not trusted_local:
        raise ValueError("MCP requires explicit trusted-local mode or configured authentication")
    loopback = host in {"127.0.0.1", "localhost", "::1"}
    trusted_container_proxy = (
        host in {"0.0.0.0", "::"}  # nosec B104
        and os.environ.get("XLSLIBERATOR_MCP_TRUSTED_CONTAINER_PROXY") == "1"
        and os.environ.get("XLSLIBERATOR_APPLICATION_CONTAINER") == "1"
        and Path("/.dockerenv").is_file()
    )
    if not loopback and not trusted_container_proxy:
        raise ValueError("trusted-local MCP may bind only to a loopback address")
    logger.info(f"Starting LibreOffice UNO MCP server on {host}:{port}")
    logger.info(f"Client endpoint: http://{host}:{port}/mcp")

    mcp.run(transport="http", host=host, port=port)


if __name__ == "__main__":
    # Run server with defaults when executed directly
    serve()
