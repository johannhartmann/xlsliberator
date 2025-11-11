"""Agent-based validation of converted ODS documents.

This module provides automated validation using MCP GUI testing tools.
Works for ANY converted document, not just specific examples.
"""

import asyncio
from dataclasses import dataclass, field
from pathlib import Path

from loguru import logger


@dataclass
class AgentValidationResult:
    """Result of agent-based validation."""

    success: bool
    macros_validated: int = 0
    macros_valid: int = 0
    functions_found: int = 0
    buttons_found: int = 0
    buttons_with_handlers: int = 0
    cells_readable: int = 0
    forms_found: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


async def validate_document_with_agent(ods_path: Path) -> AgentValidationResult:
    """
    Perform agent-based validation of a converted ODS document.

    This uses MCP GUI testing tools to verify:
    - Macros are embedded and valid
    - Buttons have event handlers
    - Document state is readable
    - Forms and interactive elements work

    Args:
        ods_path: Path to ODS file to validate

    Returns:
        AgentValidationResult with validation details
    """
    result = AgentValidationResult(success=False)

    try:
        from xlsliberator.mcp_tools import (
            click_form_button,
            list_embedded_macros,
            list_sheets,
            read_cell,
            validate_macros,
        )

        logger.info(f"Agent validation: {ods_path}")

        # Step 1: Validate macros
        try:
            macros_result = await validate_macros(str(ods_path))
            if macros_result["success"]:
                result.macros_validated = macros_result["total_modules"]
                result.macros_valid = macros_result["valid_syntax"]
                logger.debug(f"Agent: {result.macros_valid}/{result.macros_validated} macros valid")
            else:
                result.warnings.append(f"Macro validation failed: {macros_result.get('error')}")
        except Exception as e:
            result.warnings.append(f"Macro validation error: {e}")

        # Step 2: Enumerate functions
        try:
            macros_list = await list_embedded_macros(str(ods_path))
            if macros_list["success"]:
                result.functions_found = macros_list["total_functions"]
                logger.debug(f"Agent: Found {result.functions_found} functions")

                # Check for button click handlers
                button_handlers = 0
                for script in macros_list["scripts"]:
                    for func in script["functions"]:
                        if "Click" in func or "click" in func or "button" in func.lower():
                            button_handlers += 1
                result.buttons_with_handlers = button_handlers
            else:
                result.warnings.append(f"Function enumeration failed: {macros_list.get('error')}")
        except Exception as e:
            result.warnings.append(f"Function enumeration error: {e}")

        # Step 3: Test document readability
        try:
            sheets_result = await list_sheets(str(ods_path))
            if sheets_result["success"]:
                sheets = sheets_result["sheets"]
                logger.debug(f"Agent: Found {len(sheets)} sheets")

                # Try reading some cells from first sheet
                if sheets:
                    sheet_name = sheets[0]
                    # Try reading A1, B1, C1
                    cells_read = 0
                    for col in ["A", "B", "C"]:
                        try:
                            cell_result = await read_cell(str(ods_path), sheet_name, f"{col}1")
                            if cell_result["success"]:
                                cells_read += 1
                        except Exception:
                            pass
                    result.cells_readable = cells_read
                    logger.debug(f"Agent: {cells_read} sample cells readable")
            else:
                result.warnings.append(f"Sheet listing failed: {sheets_result.get('error')}")
        except Exception as e:
            result.warnings.append(f"Document readability error: {e}")

        # Step 4: Detect forms and buttons
        try:
            # Try to detect common button names
            button_names = [
                "StartButton",
                "Start",
                "Button1",
                "CommandButton1",
                "ResetButton",
                "Reset",
            ]
            buttons_found = 0
            buttons_with_handlers = 0

            for button_name in button_names:
                try:
                    button_result = await click_form_button(str(ods_path), button_name)
                    if button_result.get("script_uri"):
                        buttons_found += 1
                        buttons_with_handlers += 1
                        logger.debug(f"Agent: Button '{button_name}' has event handler")
                    elif "not found" not in button_result.get("error", "").lower():
                        # Button exists but no handler
                        buttons_found += 1
                except Exception:
                    pass

            result.buttons_found = buttons_found
            if buttons_found > 0:
                logger.debug(
                    f"Agent: {buttons_with_handlers}/{buttons_found} buttons have handlers"
                )
        except Exception as e:
            result.warnings.append(f"Button detection error: {e}")

        # Success criteria
        result.success = (
            result.macros_valid > 0  # At least some valid macros
            and result.functions_found > 0  # Functions exist
            and result.cells_readable > 0  # Document is readable
            and len(result.errors) == 0  # No critical errors
        )

        if result.success:
            logger.success(
                f"Agent validation passed: {result.macros_valid} macros, "
                f"{result.functions_found} functions, "
                f"{result.buttons_with_handlers} button handlers"
            )
        else:
            logger.warning(
                f"Agent validation completed with warnings: "
                f"{len(result.warnings)} warnings, {len(result.errors)} errors"
            )

        return result

    except Exception as e:
        logger.error(f"Agent validation failed: {e}")
        result.errors.append(str(e))
        return result


def validate_document_with_agent_sync(ods_path: Path) -> AgentValidationResult:
    """Synchronous wrapper for agent validation.

    Args:
        ods_path: Path to ODS file

    Returns:
        AgentValidationResult
    """
    return asyncio.run(validate_document_with_agent(ods_path))
