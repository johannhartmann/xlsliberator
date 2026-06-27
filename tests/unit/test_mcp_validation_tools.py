"""Tests for validation-related MCP tools."""

import asyncio
import zipfile
from pathlib import Path
from typing import Any

import openpyxl

from xlsliberator import mcp_tools
from xlsliberator.lo_worker_client import WorkerResponse


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


def test_mcp_validate_transformation_with_mocked_runner(monkeypatch: Any) -> None:
    """MCP validate tool should return certification shape."""
    from xlsliberator.certification_report import CertificationReport
    from xlsliberator.validation_models import ValidationCertification

    class DummyRunner:
        def __init__(self, _plan: object) -> None:
            pass

        def run_all(self) -> CertificationReport:
            return CertificationReport(ValidationCertification(certified=True))

    import xlsliberator.validation_runner as runner_module

    monkeypatch.setattr(runner_module, "ValidationRunner", DummyRunner)

    result = asyncio.run(mcp_tools.validate_transformation("book.xlsx"))

    assert result["success"] is True
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


def test_mcp_list_sheets_uses_worker(monkeypatch: Any) -> None:
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

    result = asyncio.run(mcp_tools.list_sheets("book.ods"))

    assert result == {"success": True, "sheets": ["Sheet1"], "count": 1}
    assert calls[0]["op"] == "list_sheets"


def test_mcp_read_cell_preserves_worker_fields(monkeypatch: Any) -> None:
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

    result = asyncio.run(mcp_tools.read_cell("book.ods", "Sheet1", "A1"))

    assert result["success"] is True
    assert result["value"] == 42.0
    assert result["formula"] == "=SUM(A1:A2)"
    assert result["type"] == "FORMULA"
