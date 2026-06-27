"""Tests for macro runtime execution through the LibreOffice worker."""

from pathlib import Path
from typing import Any

from xlsliberator.lo_worker_client import WorkerError, WorkerResponse
from xlsliberator.python_macro_manager import test_script_execution as run_script_execution


def test_script_execution_uses_worker(tmp_path: Path, monkeypatch: Any) -> None:
    """Macro execution should preserve the worker return value."""
    ods_path = tmp_path / "book.ods"
    ods_path.write_bytes(b"placeholder")
    calls = []

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, payload: dict[str, Any]) -> WorkerResponse:
            calls.append(payload)
            return WorkerResponse(
                success=True,
                op="execute_script",
                data={"executed": True, "return_value": "ok"},
                wrapper_path="/lo/python",
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    result = run_script_execution(ods_path, "vnd.sun.star.script:Module.py$main")

    assert result.success is True
    assert result.return_value == "ok"
    assert calls[0]["op"] == "execute_script"
    assert calls[0]["use_gui"] is True


def test_script_execution_preserves_runtime_error(tmp_path: Path, monkeypatch: Any) -> None:
    """Runtime script failures should be returned as script errors."""
    ods_path = tmp_path / "book.ods"
    ods_path.write_bytes(b"placeholder")

    class DummyClient:
        def __init__(self, *_args: Any, **_kwargs: Any) -> None:
            pass

        def request(self, _payload: dict[str, Any]) -> WorkerResponse:
            return WorkerResponse(
                success=False,
                op="execute_script",
                error=WorkerError(type="RuntimeError", message="boom"),
                wrapper_path="/lo/python",
            )

    import xlsliberator.lo_worker_client as worker_module

    monkeypatch.setattr(worker_module, "LibreOfficeWorkerClient", DummyClient)

    result = run_script_execution(ods_path, "vnd.sun.star.script:Module.py$main")

    assert result.success is False
    assert result.error is not None
    assert "boom" in result.error
