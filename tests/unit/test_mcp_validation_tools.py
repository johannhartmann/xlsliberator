"""Tests for validation-related MCP tools."""

import asyncio
import zipfile
from pathlib import Path
from typing import Any

import openpyxl
import pytest

from xlsliberator import mcp_tools
from xlsliberator.lo_worker_client import WorkerResponse


@pytest.fixture(autouse=True)
def configured_mcp_workspace(tmp_path: Path, monkeypatch: Any) -> None:
    """Every MCP test gets an explicit, isolated host workspace root."""

    monkeypatch.setenv("XLSLIBERATOR_WORKSPACE_ROOTS", str(tmp_path))


def test_mcp_inspect_workbook(tmp_path: Path) -> None:
    """MCP inspect tool should return JSON-serializable inventory."""
    workbook_path = tmp_path / "book.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "=1+1"
    workbook.save(workbook_path)
    workbook.close()

    result = asyncio.run(mcp_tools.inspect_workbook(str(workbook_path)))

    assert result["success"] is True
    assert result["inventory"]["formulas"][0]["formula_text"] == "=1+1"
    assert result["inventory"]["schema_version"] == "3.0.0"
    assert result["inventory"]["artifacts"]
    assert result["inventory"]["metadata"]["canonical_inventory"]["artifact_count"] > 0


def test_mcp_rejects_host_path_outside_configured_workspace(tmp_path: Path) -> None:
    outside = tmp_path.parent / "outside.xlsx"
    outside.write_bytes(b"host secret")

    result = asyncio.run(mcp_tools.inspect_workbook(str(outside)))

    assert result["success"] is False
    assert result["error"]["type"] == "WorkspaceAccessError"


def test_mcp_conversion_success_follows_report(tmp_path: Path, monkeypatch: Any) -> None:
    """Transport success must not turn a failed conversion into success."""
    from xlsliberator.report import ConversionReport

    monkeypatch.setattr(
        mcp_tools,
        "convert_api",
        lambda **_kwargs: ConversionReport(
            input_file="in.xlsx",
            output_file="out.ods",
            success=False,
            errors=["conversion failed"],
        ),
    )

    source = tmp_path / "in.xlsx"
    source.write_bytes(b"fixture")
    result = asyncio.run(mcp_tools.convert_excel_to_ods(str(source), str(tmp_path / "out.ods")))

    assert result["transport_success"] is True
    assert result["success"] is False
    assert result["operation_status"] == "failed"


def test_unimplemented_keyboard_input_is_unavailable() -> None:
    """The keyboard placeholder must never report fabricated success."""
    result = asyncio.run(mcp_tools.send_keyboard_input("book.ods", ["ENTER"]))

    assert result["success"] is False
    assert result["implemented"] is False
    assert result["operation_status"] == "unavailable"
    assert result["keys_sent"] == 0


def test_mcp_validate_transformation_with_mocked_runner(tmp_path: Path, monkeypatch: Any) -> None:
    """MCP validate tool should return certification shape."""
    from xlsliberator.certification_report import CertificationReport
    from xlsliberator.validation_models import ValidationCertification, ValidationGateResult

    class DummyRunner:
        def __init__(self, _plan: object) -> None:
            pass

        def run_all(self) -> CertificationReport:
            return CertificationReport(
                ValidationCertification(
                    gate_results=[
                        ValidationGateResult(gate_name="runtime", passed=True, message="ok")
                    ]
                )
            )

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "ValidationRunner", DummyRunner)

    source = tmp_path / "book.xlsx"
    source.write_bytes(b"fixture")
    result = asyncio.run(mcp_tools.validate_transformation(str(source)))

    assert result["success"] is True
    assert result["transport_success"] is True
    assert result["certification"]["certified"] is True


def test_mcp_list_controls_and_event_bindings(tmp_path: Path) -> None:
    """MCP control listing tools should return serializable data."""
    ods_path = tmp_path / "controls.ods"
    content = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
 <office:body><office:spreadsheet><table:table table:name="Sheet1">
  <form:button form:id="btn1" form:name="StartButton">
   <form:event-listener form:event-name="approveAction"
    xlink:href="vnd.sun.star.script:Module.py$Start?language=Python&amp;location=document"/>
  </form:button>
 </table:table></office:spreadsheet></office:body>
</office:document-content>"""
    with zipfile.ZipFile(ods_path, "w") as archive:
        archive.writestr("content.xml", content)

    controls = asyncio.run(mcp_tools.list_controls(str(ods_path)))
    events = asyncio.run(mcp_tools.list_event_bindings(str(ods_path)))

    assert controls["success"] is True
    assert controls["count"] == 1
    assert events["success"] is True
    assert events["count"] == 1


def test_mcp_list_sheets_uses_worker(tmp_path: Path, monkeypatch: Any) -> None:
    """UNO-backed MCP tools should route through the worker client."""
    calls = []

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, payload: dict[str, Any]) -> WorkerResponse:
            calls.append(payload)
            return WorkerResponse(
                success=True,
                op="list_sheets",
                data={"sheets": ["Sheet1"], "count": 1},
                wrapper_path="/lo/python",
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    source = tmp_path / "book.ods"
    source.write_bytes(b"fixture")
    result = asyncio.run(mcp_tools.list_sheets(str(source)))

    assert result["success"] is True
    assert result["transport_success"] is True
    assert result["operation_status"] == "passed"
    assert result["sheets"] == ["Sheet1"]
    assert result["count"] == 1
    assert calls[0]["op"] == "list_sheets"


def test_mcp_read_cell_preserves_worker_fields(tmp_path: Path, monkeypatch: Any) -> None:
    """Cell reads should preserve value, formula, and type from the worker."""

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, _payload: dict[str, Any]) -> WorkerResponse:
            return WorkerResponse(
                success=True,
                op="read_cell",
                data={"value": 42.0, "formula": "=SUM(A1:A2)", "type": "FORMULA"},
                wrapper_path="/lo/python",
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    source = tmp_path / "book.ods"
    source.write_bytes(b"fixture")
    result = asyncio.run(mcp_tools.read_cell(str(source), "Sheet1", "A1"))

    assert result["success"] is True
    assert result["value"] == 42.0
    assert result["formula"] == "=SUM(A1:A2)"
    assert result["type"] == "FORMULA"


def test_missing_docker_sandbox_makes_required_macro_execution_unavailable(
    tmp_path: Path, monkeypatch: Any
) -> None:
    from xlsliberator.lo_worker_client import WorkerError

    class UnavailableClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, _payload: dict[str, Any]) -> WorkerResponse:
            return WorkerResponse(
                success=False,
                op="execute_script",
                error=WorkerError(
                    type="DockerRuntimeUnavailable",
                    message="Docker sandbox is unavailable",
                ),
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", UnavailableClient)
    workbook = tmp_path / "macro.ods"
    workbook.write_bytes(b"fixture")

    result = asyncio.run(
        mcp_tools.test_macro_execution(
            str(workbook),
            "vnd.sun.star.script:Module.py$run?language=Python&location=document",
        )
    )

    assert result["success"] is False
    assert result["operation_status"] == "unavailable"
    assert result["capability_available"] is False


def test_missing_control_is_failed_handler_action(tmp_path: Path, monkeypatch: Any) -> None:
    """A handler request for a missing discovered control must fail."""

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, _payload: dict[str, Any]) -> WorkerResponse:
            from xlsliberator.lo_worker_client import WorkerError

            return WorkerResponse(
                success=False,
                op="execute_button_handler",
                error=WorkerError(type="ValueError", message="Button 'Missing' not found"),
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    source = tmp_path / "book.ods"
    source.write_bytes(b"fixture")
    result = asyncio.run(mcp_tools.execute_button_handler(str(source), "Missing"))

    assert result["success"] is False
    assert result["operation_status"] == "failed"
    assert result["implemented"] is True
    assert result["error"]["type"] == "ValueError"


def test_direct_handler_invocation_is_labelled_accurately(tmp_path: Path, monkeypatch: Any) -> None:
    """Direct script invocation must never be represented as a GUI click."""
    calls: list[dict[str, Any]] = []

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, payload: dict[str, Any]) -> WorkerResponse:
            calls.append(payload)
            return WorkerResponse(
                success=True,
                op="execute_button_handler",
                data={"handler_executed": True, "script_uri": "Module.py$run"},
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    source = tmp_path / "book.ods"
    source.write_bytes(b"fixture")
    direct = asyncio.run(mcp_tools.execute_button_handler(str(source), "Run"))
    click = asyncio.run(mcp_tools.click_form_button(str(source), "Run"))

    assert calls[0]["op"] == "execute_button_handler"
    assert direct["success"] is True
    assert direct["handler_executed"] is True
    assert "clicked" not in direct
    assert click["success"] is False
    assert click["implemented"] is False
    assert click["operation_status"] == "unavailable"
