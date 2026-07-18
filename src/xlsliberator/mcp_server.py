"""FastMCP server for stateful LibreOffice-only runtime sessions.

Usage:
    xlsliberator libreoffice-mcp-serve --port 8000
"""

import os
from pathlib import Path

from fastmcp import FastMCP
from loguru import logger

from xlsliberator.container_boundary import require_application_container
from xlsliberator.libreoffice_mcp import (
    build_interactive_game_target,
    bundle_interactive_game_replays,
    capture_screenshot,
    close,
    collect_logs,
    create_session,
    destroy_session,
    dispatch_control_event,
    execute_python_macro,
    export_pdf,
    inspect_document,
    list_controls,
    list_formulas,
    list_sheets,
    open_document,
    read_cells,
    recalculate,
    reopen,
    run_interactive_game_scenario,
    save,
    send_keyboard_event,
    write_cells,
)

mcp = FastMCP(name="XLSLiberator LibreOffice Runtime")

for tool in (
    create_session,
    open_document,
    inspect_document,
    list_sheets,
    read_cells,
    write_cells,
    list_formulas,
    recalculate,
    list_controls,
    dispatch_control_event,
    send_keyboard_event,
    execute_python_macro,
    capture_screenshot,
    export_pdf,
    save,
    close,
    reopen,
    collect_logs,
    destroy_session,
    build_interactive_game_target,
    run_interactive_game_scenario,
    bundle_interactive_game_replays,
):
    mcp.tool(tool)


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
    logger.info(f"Starting stateful LibreOffice runtime MCP server on {host}:{port}")
    logger.info(f"Client endpoint: http://{host}:{port}/mcp")

    mcp.run(transport="http", host=host, port=port)


if __name__ == "__main__":
    # Run server with defaults when executed directly
    serve()
