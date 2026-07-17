"""Tests for the Docker-only LibreOffice worker client."""

from __future__ import annotations

from typing import Any

import pytest

from xlsliberator.docker_runtime import (
    DockerRuntimeTimeout,
    DockerRuntimeUnavailable,
    MalformedWorkerResponse,
)
from xlsliberator.lo_worker_client import (
    LibreOfficeWorkerClient,
    discover_libreoffice_python_wrapper,
)


class FakeRuntime:
    def __init__(self, response: dict[str, Any] | None = None, error: Exception | None = None):
        self.response = response or {}
        self.error = error
        self.requests: list[dict[str, Any]] = []

    def request(self, payload: dict[str, Any]) -> dict[str, Any]:
        self.requests.append(payload)
        if self.error:
            raise self.error
        return self.response


def test_host_wrapper_discovery_is_disabled() -> None:
    assert discover_libreoffice_python_wrapper("/Applications/LibreOffice/soffice") is None


def test_client_maps_valid_docker_response() -> None:
    runtime = FakeRuntime(
        {"success": True, "op": "ping", "data": {"uno_importable": True}, "error": None}
    )

    response = LibreOfficeWorkerClient(runtime=runtime).ping()  # type: ignore[arg-type]

    assert response.success is True
    assert response.data["uno_importable"] is True
    assert runtime.requests[0]["op"] == "ping"
    assert runtime.requests[0]["timeout_seconds"] == 10


def test_client_preserves_worker_error() -> None:
    runtime = FakeRuntime(
        {
            "success": False,
            "op": "ping",
            "data": {},
            "error": {"type": "ImportError", "message": "No module named uno"},
        }
    )

    response = LibreOfficeWorkerClient(runtime=runtime).ping()  # type: ignore[arg-type]

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "ImportError"


def test_client_fails_closed_when_docker_is_unavailable() -> None:
    runtime = FakeRuntime(error=DockerRuntimeUnavailable("docker missing"))

    response = LibreOfficeWorkerClient(runtime=runtime).ping()  # type: ignore[arg-type]

    assert response.success is False
    assert response.error is not None
    assert response.error.type == "DockerRuntimeUnavailable"
    assert "host fallback is disabled" in response.error.message


def test_legacy_host_arguments_cannot_select_local_runtime() -> None:
    runtime = FakeRuntime({"success": True, "op": "ping", "data": {}, "error": None})
    client = LibreOfficeWorkerClient(
        office_executable="/usr/bin/soffice",
        python_wrapper="/usr/bin/python3",
        runtime=runtime,  # type: ignore[arg-type]
    )

    assert client.office_executable is None
    assert client.python_wrapper is None
    assert client.ping().success is True


@pytest.mark.parametrize(
    ("error", "expected_type"),
    [
        (DockerRuntimeTimeout("job exceeded limit"), "DockerRuntimeTimeout"),
        (MalformedWorkerResponse("bad json"), "MalformedWorkerResponse"),
    ],
)
def test_client_does_not_hide_runtime_protocol_failures(
    error: Exception, expected_type: str
) -> None:
    response = LibreOfficeWorkerClient(runtime=FakeRuntime(error=error)).ping()  # type: ignore[arg-type]

    assert response.success is False
    assert response.error is not None
    assert response.error.type == expected_type
