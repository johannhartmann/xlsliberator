"""Tests for data-driven agent validation behavior."""

import asyncio
from pathlib import Path
from typing import Any

from xlsliberator.agent_validator import validate_document_with_agent


async def _ok_validate_macros(_ods_path: str) -> dict[str, Any]:
    return {"success": True, "total_modules": 0, "valid_syntax": 0}


async def _ok_list_embedded_macros(_ods_path: str) -> dict[str, Any]:
    return {"success": True, "scripts": [], "total_functions": 0}


async def _ok_runtime(_ods_path: str) -> dict[str, Any]:
    return {"success": True, "operation_status": "passed", "stages": {"open": {"status": "passed"}}}


async def _failed_runtime(_ods_path: str) -> dict[str, Any]:
    return {"success": False, "operation_status": "failed"}


async def _invalid_validate_macros(_ods_path: str) -> dict[str, Any]:
    return {"success": True, "total_modules": 2, "valid_syntax": 1}


def test_agent_validator_no_macros_controls_can_pass(monkeypatch: Any, tmp_path: Path) -> None:
    """A workbook without macros or controls can pass with full runtime evidence."""
    import xlsliberator.control_inventory as inventory_module
    import xlsliberator.mcp_tools as mcp_module

    monkeypatch.setattr(inventory_module, "extract_controls_from_ods", lambda _path: [])
    monkeypatch.setattr(inventory_module, "extract_event_bindings_from_ods", lambda _path: [])
    monkeypatch.setattr(mcp_module, "validate_macros", _ok_validate_macros)
    monkeypatch.setattr(mcp_module, "list_embedded_macros", _ok_list_embedded_macros)
    monkeypatch.setattr(mcp_module, "validate_document_runtime", _ok_runtime)

    result = asyncio.run(validate_document_with_agent(tmp_path / "book.ods"))

    assert result.success
    assert result.buttons_found == 0
    assert result.cells_readable == 0
    assert result.runtime_status == "passed"


def test_agent_validator_fails_when_embedded_macros_invalid(
    monkeypatch: Any, tmp_path: Path
) -> None:
    """Embedded macros that fail syntax validation must fail agent validation."""
    import xlsliberator.control_inventory as inventory_module
    import xlsliberator.mcp_tools as mcp_module

    monkeypatch.setattr(inventory_module, "extract_controls_from_ods", lambda _path: [])
    monkeypatch.setattr(inventory_module, "extract_event_bindings_from_ods", lambda _path: [])
    monkeypatch.setattr(mcp_module, "validate_macros", _invalid_validate_macros)
    monkeypatch.setattr(mcp_module, "list_embedded_macros", _ok_list_embedded_macros)
    monkeypatch.setattr(mcp_module, "validate_document_runtime", _ok_runtime)

    result = asyncio.run(validate_document_with_agent(tmp_path / "book.ods"))

    assert not result.success
    assert result.macros_validated == 2
    assert result.macros_valid == 1


def test_agent_validator_cannot_pass_from_a1_c1_reads(monkeypatch: Any, tmp_path: Path) -> None:
    """Sample cell readability cannot replace complete target-runtime evidence."""
    import xlsliberator.control_inventory as inventory_module
    import xlsliberator.mcp_tools as mcp_module

    monkeypatch.setattr(inventory_module, "extract_controls_from_ods", lambda _path: [])
    monkeypatch.setattr(inventory_module, "extract_event_bindings_from_ods", lambda _path: [])
    monkeypatch.setattr(mcp_module, "validate_macros", _ok_validate_macros)
    monkeypatch.setattr(mcp_module, "list_embedded_macros", _ok_list_embedded_macros)
    monkeypatch.setattr(mcp_module, "validate_document_runtime", _failed_runtime)

    result = asyncio.run(validate_document_with_agent(tmp_path / "book.ods"))

    assert result.success is False
    assert result.cells_readable == 0
    assert result.runtime_status == "failed"
