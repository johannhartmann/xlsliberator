"""Client for running UNO operations in LibreOffice's Python wrapper."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
from contextlib import suppress
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

DEFAULT_WORKER_TIMEOUT_SECONDS = 30
UNO_WORKER_UNAVAILABLE = "UNO worker unavailable"


@dataclass(frozen=True)
class WorkerError:
    """Structured worker error returned to callers."""

    type: str
    message: str
    traceback: str | None = None
    stderr: str | None = None
    returncode: int | None = None
    wrapper_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "type": self.type,
            "message": self.message,
            "traceback": self.traceback,
            "stderr": self.stderr,
            "returncode": self.returncode,
            "wrapper_path": self.wrapper_path,
        }


@dataclass(frozen=True)
class WorkerResponse:
    """Response from the out-of-process LibreOffice worker."""

    success: bool
    op: str
    data: dict[str, Any] = field(default_factory=dict)
    error: WorkerError | None = None
    wrapper_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "success": self.success,
            "op": self.op,
            "data": self.data,
            "error": self.error.to_dict() if self.error else None,
            "wrapper_path": self.wrapper_path,
        }


class LibreOfficeWorkerClient:
    """Run one UNO worker request per LibreOffice Python process."""

    def __init__(
        self,
        office_executable: str | None = None,
        python_wrapper: str | None = None,
        timeout_seconds: int = DEFAULT_WORKER_TIMEOUT_SECONDS,
    ) -> None:
        self.office_executable = office_executable or _discover_office_executable()
        self.python_wrapper = python_wrapper or discover_libreoffice_python_wrapper(
            self.office_executable
        )
        self.timeout_seconds = timeout_seconds

    def ping(self) -> WorkerResponse:
        """Verify the wrapper can import pyuno."""
        return self.request({"op": "ping"}, timeout_seconds=10)

    def request(
        self,
        payload: dict[str, Any],
        timeout_seconds: int | None = None,
    ) -> WorkerResponse:
        """Run a single worker request and return a structured response."""
        op = str(payload.get("op", "unknown"))
        if not self.python_wrapper:
            return self._unavailable(op, "LibreOffice Python wrapper was not found")
        if not Path(self.python_wrapper).exists():
            return self._unavailable(op, f"LibreOffice Python wrapper not found: {self.python_wrapper}")

        request_payload = dict(payload)
        if self.office_executable:
            request_payload.setdefault("office_executable", self.office_executable)

        env = dict(os.environ)
        source_root = str(Path(__file__).resolve().parents[1])
        existing_pythonpath = env.get("PYTHONPATH")
        env["PYTHONPATH"] = (
            source_root
            if not existing_pythonpath
            else os.pathsep.join([source_root, existing_pythonpath])
        )

        cmd = [self.python_wrapper, "-m", "xlsliberator.lo_worker"]
        timeout = timeout_seconds or request_payload.get("timeout_seconds") or self.timeout_seconds
        try:
            result = subprocess.run(
                cmd,
                input=json.dumps(request_payload),
                capture_output=True,
                text=True,
                timeout=float(timeout),
                check=False,
                env=env,
            )
        except subprocess.TimeoutExpired as exc:
            return self._error_response(
                op,
                WorkerError(
                    type="TimeoutExpired",
                    message=f"LibreOffice worker timed out after {timeout} seconds",
                    stderr=_bytes_or_text(exc.stderr),
                    wrapper_path=self.python_wrapper,
                ),
            )
        except OSError as exc:
            return self._error_response(
                op,
                WorkerError(
                    type=type(exc).__name__,
                    message=str(exc),
                    wrapper_path=self.python_wrapper,
                ),
        )

        stdout = result.stdout.strip()
        if not stdout and result.returncode != 0:
            return self._error_response(
                op,
                WorkerError(
                    type="WorkerNonZeroExit",
                    message=f"Worker exited with status {result.returncode}",
                    stderr=result.stderr,
                    returncode=result.returncode,
                    wrapper_path=self.python_wrapper,
                ),
            )
        try:
            raw_response = json.loads(stdout)
        except json.JSONDecodeError as exc:
            return self._error_response(
                op,
                WorkerError(
                    type="MalformedWorkerJSON",
                    message=f"Worker did not return valid JSON: {exc}",
                    stderr=result.stderr,
                    returncode=result.returncode,
                    wrapper_path=self.python_wrapper,
                ),
            )

        error = None
        raw_error = raw_response.get("error")
        if raw_error:
            error = WorkerError(
                type=str(raw_error.get("type", "WorkerError")),
                message=str(raw_error.get("message", "")),
                traceback=raw_error.get("traceback"),
                stderr=result.stderr or raw_error.get("stderr"),
                returncode=result.returncode,
                wrapper_path=self.python_wrapper,
            )
        elif result.returncode != 0:
            error = WorkerError(
                type="WorkerNonZeroExit",
                message=f"Worker exited with status {result.returncode}",
                stderr=result.stderr,
                returncode=result.returncode,
                wrapper_path=self.python_wrapper,
            )

        success = bool(raw_response.get("success")) and error is None
        return WorkerResponse(
            success=success,
            op=str(raw_response.get("op", op)),
            data=dict(raw_response.get("data") or {}),
            error=error,
            wrapper_path=self.python_wrapper,
        )

    def _unavailable(self, op: str, message: str) -> WorkerResponse:
        return self._error_response(
            op,
            WorkerError(
                type="UnoWorkerUnavailable",
                message=f"{UNO_WORKER_UNAVAILABLE}: {message}",
                wrapper_path=self.python_wrapper,
            ),
        )

    def _error_response(self, op: str, error: WorkerError) -> WorkerResponse:
        return WorkerResponse(
            success=False,
            op=op,
            data={},
            error=error,
            wrapper_path=self.python_wrapper,
        )


def check_worker(
    office_executable: str | None = None,
    python_wrapper: str | None = None,
) -> WorkerResponse:
    """Return the worker ping result without raising on local setup problems."""
    return LibreOfficeWorkerClient(
        office_executable=office_executable,
        python_wrapper=python_wrapper,
    ).ping()


def worker_unavailable_message(response: WorkerResponse) -> str:
    """Format a compact worker failure message for tool responses."""
    if response.error is None:
        return UNO_WORKER_UNAVAILABLE
    wrapper = f" (wrapper: {response.wrapper_path})" if response.wrapper_path else ""
    stderr = f"; stderr: {response.error.stderr.strip()}" if response.error.stderr else ""
    if response.error.type in {
        "UnoWorkerUnavailable",
        "ImportError",
        "MalformedWorkerJSON",
        "TimeoutExpired",
        "WorkerNonZeroExit",
    }:
        return f"{UNO_WORKER_UNAVAILABLE}: {response.error.message}{wrapper}{stderr}"
    return f"UNO worker error: {response.error.message}{wrapper}{stderr}"


def discover_libreoffice_python_wrapper(office_executable: str | None = None) -> str | None:
    """Discover LibreOffice's Python wrapper for pyuno execution."""
    candidates: list[Path] = []

    if office_executable:
        executable_path = Path(office_executable)
        with_resolved = [executable_path]
        with suppress(OSError):
            with_resolved.append(executable_path.resolve())
        for path in with_resolved:
            candidates.extend(_wrapper_candidates_from_executable(path))

    for app_path in (
        Path("/Applications/LibreOffice.app"),
        Path.home() / "Applications/LibreOffice.app",
    ):
        candidates.extend(
            [
                app_path / "Contents/Resources/python",
                app_path / "Contents/MacOS/python",
            ]
        )

    for path_text in (
        "/usr/lib/libreoffice/program/python",
        "/usr/lib64/libreoffice/program/python",
        "/opt/libreoffice/program/python",
        "/snap/libreoffice/current/lib/libreoffice/program/python",
    ):
        candidates.append(Path(path_text))

    seen: set[Path] = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        if candidate.exists() and os.access(candidate, os.X_OK):
            return str(candidate)
    return None


def _wrapper_candidates_from_executable(executable: Path) -> list[Path]:
    candidates: list[Path] = []
    parts = executable.parts
    if "Contents" in parts:
        contents_index = parts.index("Contents")
        contents_dir = Path(*parts[: contents_index + 1])
        candidates.extend(
            [
                contents_dir / "Resources/python",
                contents_dir / "MacOS/python",
            ]
        )

    parent = executable.parent
    candidates.extend(
        [
            parent / "python",
            parent.parent / "Resources/python",
            parent.parent / "MacOS/python",
        ]
    )
    return candidates


def _discover_office_executable() -> str | None:
    for executable_name in ("soffice", "libreoffice"):
        executable = shutil.which(executable_name)
        if executable:
            return executable
    macos_soffice = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if macos_soffice.exists():
        return str(macos_soffice)
    return None


def _bytes_or_text(value: bytes | str | None) -> str | None:
    if value is None:
        return None
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="replace")
    return value
