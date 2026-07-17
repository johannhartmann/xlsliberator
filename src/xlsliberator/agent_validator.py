"""Evidence-driven validation of converted ODS documents."""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from loguru import logger


@dataclass
class AgentValidationResult:
    """Result derived from inventories and deterministic Docker runtime evidence."""

    success: bool
    macros_validated: int = 0
    macros_valid: int = 0
    functions_found: int = 0
    buttons_found: int = 0
    buttons_with_handlers: int = 0
    cells_readable: int = 0
    forms_found: int = 0
    runtime_status: str = "not_run"
    evidence: dict[str, Any] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


async def validate_document_with_agent(ods_path: Path) -> AgentValidationResult:
    """Validate package inventory and complete target-runtime lifecycle evidence.

    This compatibility entry point no longer performs heuristic GUI probes or
    samples A1-C1. It cannot pass without a successful open/recalculate/save/
    close/reopen/package run in the pinned disposable Docker target.
    """
    from xlsliberator.control_inventory import (
        extract_controls_from_ods,
        extract_event_bindings_from_ods,
    )
    from xlsliberator.mcp_tools import (
        list_embedded_macros,
        validate_document_runtime,
        validate_macros,
    )

    result = AgentValidationResult(success=False)
    logger.info(f"Evidence-driven validation: {ods_path}")

    try:
        controls = extract_controls_from_ods(ods_path)
        bindings = extract_event_bindings_from_ods(ods_path)
        buttons = [control for control in controls if "button" in control.control_type.lower()]
        bound_control_ids = {
            binding.control_id
            for binding in bindings
            if binding.control_id and binding.target_script_uri
        }
        result.buttons_found = len(buttons)
        result.buttons_with_handlers = sum(
            1
            for button in buttons
            if button.id in bound_control_ids or button.name in bound_control_ids
        )
        result.forms_found = len({control.sheet for control in controls if control.sheet})
        result.evidence["inventory"] = {
            "controls": [control.model_dump(mode="json") for control in controls],
            "event_bindings": [binding.model_dump(mode="json") for binding in bindings],
        }
        unbound = sorted(
            button.name
            for button in buttons
            if button.id not in bound_control_ids and button.name not in bound_control_ids
        )
        if unbound:
            result.errors.append(f"Buttons without discovered target handlers: {unbound}")
    except Exception as exc:
        result.errors.append(f"Control inventory failed: {exc}")

    macro_result = await validate_macros(str(ods_path))
    result.evidence["macro_validation"] = macro_result
    result.macros_validated = int(macro_result.get("total_modules", 0))
    result.macros_valid = int(macro_result.get("valid_syntax", 0))
    macros_complete = (
        result.macros_validated == result.macros_valid
        and int(macro_result.get("syntax_errors", 0)) == 0
        and int(macro_result.get("missing_exported_scripts", 0)) == 0
    )
    if not macro_result.get("success") or not macros_complete:
        result.errors.append("Embedded macro validation did not pass")

    macro_inventory = await list_embedded_macros(str(ods_path))
    result.evidence["macro_inventory"] = macro_inventory
    result.functions_found = int(macro_inventory.get("total_functions", 0))
    if not macro_inventory.get("success"):
        result.errors.append("Embedded macro inventory failed")

    runtime_result = await validate_document_runtime(str(ods_path))
    result.evidence["target_runtime"] = runtime_result
    result.runtime_status = str(runtime_result.get("operation_status", "failed"))
    if not runtime_result.get("success"):
        result.errors.append("Pinned Docker target runtime validation did not pass")

    result.success = not result.errors
    if result.success:
        logger.success("Evidence-driven validation passed")
    else:
        logger.warning(f"Evidence-driven validation failed with {len(result.errors)} error(s)")
    return result


def validate_document_with_agent_sync(ods_path: Path) -> AgentValidationResult:
    """Synchronous wrapper for the evidence-driven validator."""
    return asyncio.run(validate_document_with_agent(ods_path))
