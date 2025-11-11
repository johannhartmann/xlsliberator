"""FastMCP 2.0 server for LibreOffice UNO operations.

Exposes xlsliberator functionality through Model Context Protocol (MCP)
for integration with Claude Agent SDK and other MCP clients.

Usage:
    # Run with HTTP streaming transport
    python -m xlsliberator.mcp_server

    # Or via CLI
    xlsliberator mcp-serve --port 8000

    # Connect client to: http://localhost:8000/mcp
"""

from fastmcp import FastMCP
from loguru import logger

from xlsliberator.mcp_tools import (
    click_form_button,
    compare_formulas,
    convert_excel_to_ods,
    embed_macros,
    get_cell_colors,
    get_sheet_data,
    list_embedded_macros,
    list_sheets,
    open_document_gui,
    read_cell,
    recalculate_document,
    send_keyboard_input,
    take_screenshot,
    test_macro_execution,
    validate_macros,
)

# Create FastMCP server instance
mcp = FastMCP(name="LibreOffice UNO")


# ==============================================================================
# Register Tools
# ==============================================================================

# Document Operations
mcp.tool(convert_excel_to_ods)
mcp.tool(recalculate_document)

# Cell and Sheet Operations
mcp.tool(read_cell)
mcp.tool(list_sheets)
mcp.tool(get_sheet_data)

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
mcp.tool(send_keyboard_input)
mcp.tool(get_cell_colors)
mcp.tool(take_screenshot)


def serve(host: str = "0.0.0.0", port: int = 8000) -> None:  # nosec B104
    """Start the MCP server with HTTP streaming transport.

    Note: Binds to 0.0.0.0 by design for Docker/container environments.

    Args:
        host: Host address to bind to (default: 0.0.0.0)
        port: Port number (default: 8000)
    """
    logger.info(f"Starting LibreOffice UNO MCP server on {host}:{port}")
    logger.info(f"Client endpoint: http://{host}:{port}/mcp")

    mcp.run(transport="http", host=host, port=port)


if __name__ == "__main__":
    # Run server with defaults when executed directly
    serve()
