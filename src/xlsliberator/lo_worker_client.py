"""Client for UNO operations in the authoritative LibreOffice Docker runtime."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from xlsliberator.docker_runtime import DockerRuntimeUnavailable, LibreOfficeDockerRuntime

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
    """Response from one disposable LibreOffice container."""

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
    """Send each UNO request to a new container; never inspect or run host office."""

    def __init__(
        self,
        office_executable: str | None = None,
        python_wrapper: str | None = None,
        timeout_seconds: int = DEFAULT_WORKER_TIMEOUT_SECONDS,
        *,
        runtime: LibreOfficeDockerRuntime | None = None,
    ) -> None:
        # Retained only so old callers fail safely instead of silently selecting a host runtime.
        del office_executable, python_wrapper
        self.timeout_seconds = timeout_seconds
        self.runtime = runtime or LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds)
        self.python_wrapper: str | None = None
        self.office_executable: str | None = None

    def ping(self) -> WorkerResponse:
        """Verify that the container runtime has its matching PyUNO bridge."""
        return self.request({"op": "ping"}, timeout_seconds=10)

    def request(
        self,
        payload: dict[str, Any],
        timeout_seconds: int | None = None,
    ) -> WorkerResponse:
        """Run a single Docker worker request and return a structured response."""
        op = str(payload.get("op", "unknown"))
        request_payload = dict(payload)
        request_payload["timeout_seconds"] = (
            timeout_seconds or request_payload.get("timeout_seconds") or self.timeout_seconds
        )
        try:
            raw = self.runtime.request(request_payload)
        except (DockerRuntimeUnavailable, OSError) as exc:
            return self._error_response(
                op,
                WorkerError(
                    type=type(exc).__name__,
                    message=f"{UNO_WORKER_UNAVAILABLE}: {exc}; host fallback is disabled",
                ),
            )

        raw_error = raw.get("error") or None
        error = None
        if raw_error:
            error = WorkerError(
                type=str(raw_error.get("type", "WorkerError")),
                message=str(raw_error.get("message", "")),
                traceback=raw_error.get("traceback"),
                stderr=raw_error.get("stderr"),
                returncode=raw_error.get("returncode"),
            )
        return WorkerResponse(
            success=bool(raw.get("success")) and error is None,
            op=str(raw.get("op", op)),
            data=dict(raw.get("data") or {}),
            error=error,
            wrapper_path=None,
        )

    def _error_response(self, op: str, error: WorkerError) -> WorkerResponse:
        return WorkerResponse(success=False, op=op, data={}, error=error, wrapper_path=None)


def check_worker(
    office_executable: str | None = None,
    python_wrapper: str | None = None,
) -> WorkerResponse:
    """Return the Docker worker ping result without raising setup problems."""
    return LibreOfficeWorkerClient(
        office_executable=office_executable,
        python_wrapper=python_wrapper,
    ).ping()


def worker_unavailable_message(response: WorkerResponse) -> str:
    """Format a compact worker failure message for tool responses."""
    if response.error is None:
        return UNO_WORKER_UNAVAILABLE
    if response.error.type in {
        "DockerRuntimeUnavailable",
        "DockerRuntimeTimeout",
        "ImportError",
        "MalformedWorkerJSON",
        "MalformedWorkerResponse",
        "TimeoutExpired",
        "WorkerNonZeroExit",
    }:
        return f"{UNO_WORKER_UNAVAILABLE}: {response.error.message}"
    return f"UNO worker error: {response.error.message}"


def discover_libreoffice_python_wrapper(office_executable: str | None = None) -> None:
    """Host wrapper discovery is intentionally disabled."""
    del office_executable
    return None


def _discover_office_executable() -> None:
    """Host LibreOffice discovery is intentionally disabled."""
    return None
