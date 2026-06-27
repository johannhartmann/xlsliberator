"""Tests for the out-of-process LibreOffice worker client."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any

from xlsliberator.lo_worker_client import (
    LibreOfficeWorkerClient,
    discover_libreoffice_python_wrapper,
)


def test_discovers_macos_wrapper_from_backend(tmp_path: Path, monkeypatch: Any) -> None:
    """The client should map a macOS soffice executable to Resources/python."""
    app = tmp_path / "LibreOffice.app"
    soffice = app / "Contents/MacOS/soffice"
    wrapper = app / "Contents/Resources/python"
    soffice.parent.mkdir(parents=True)
    wrapper.parent.mkdir(parents=True)
    soffice.write_text("#!/bin/sh\n")
    wrapper.write_text("#!/bin/sh\n")
    wrapper.chmod(0o755)

    monkeypatch.setattr(Path, "home", lambda: tmp_path)

    assert discover_libreoffice_python_wrapper(str(soffice)) == str(wrapper)


def test_client_handles_valid_json(monkeypatch: Any) -> None:
    """Successful worker JSON should be converted into a typed response."""

    def fake_run(*_args: Any, **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=["python"],
            returncode=0,
            stdout=json.dumps(
                {"success": True, "op": "ping", "data": {"uno_importable": True}, "error": None}
            ),
            stderr="",
        )

    monkeypatch.setattr(subprocess, "run", fake_run)

    response = LibreOfficeWorkerClient(python_wrapper="/bin/sh").ping()

    assert response.success is True
    assert response.data["uno_importable"] is True


def test_client_handles_malformed_json(monkeypatch: Any) -> None:
    """Non-JSON stdout should become a structured client error."""

    def fake_run(*_args: Any, **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=["python"],
            returncode=0,
            stdout="not json",
            stderr="debug",
        )

    monkeypatch.setattr(subprocess, "run", fake_run)

    response = LibreOfficeWorkerClient(python_wrapper="/bin/sh").ping()

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "MalformedWorkerJSON"
    assert response.error.stderr == "debug"


def test_client_handles_timeout(monkeypatch: Any) -> None:
    """Worker timeouts should not raise into callers."""

    def fake_run(*_args: Any, **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        raise subprocess.TimeoutExpired(cmd=["python"], timeout=1, stderr="slow")

    monkeypatch.setattr(subprocess, "run", fake_run)

    response = LibreOfficeWorkerClient(python_wrapper="/bin/sh").ping()

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "TimeoutExpired"


def test_client_handles_nonzero_exit_with_json_error(monkeypatch: Any) -> None:
    """Worker JSON failures should preserve worker error fields."""

    def fake_run(*_args: Any, **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(
            args=["python"],
            returncode=1,
            stdout=json.dumps(
                {
                    "success": False,
                    "op": "ping",
                    "data": {},
                    "error": {
                        "type": "ImportError",
                        "message": "No module named uno",
                        "traceback": "trace",
                    },
                }
            ),
            stderr="stderr",
        )

    monkeypatch.setattr(subprocess, "run", fake_run)

    response = LibreOfficeWorkerClient(python_wrapper="/bin/sh").ping()

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "ImportError"
    assert response.error.stderr == "stderr"


def test_client_handles_unavailable_wrapper() -> None:
    """Missing wrappers should produce an unavailable response."""
    response = LibreOfficeWorkerClient(python_wrapper="/missing/libreoffice/python").ping()

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "UnoWorkerUnavailable"
