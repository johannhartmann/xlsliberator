"""End-to-end smoke test for the trusted web Docker orchestrator."""

from __future__ import annotations

import json
import os
import shutil
import socket
import subprocess
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path

import openpyxl
import pytest


def _docker(*arguments: str, timeout: int = 600) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["docker", *arguments],
        check=True,
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def _get_json(url: str) -> dict[str, object]:
    with urllib.request.urlopen(url, timeout=15) as response:  # noqa: S310
        payload = json.loads(response.read())
    assert isinstance(payload, dict)
    return payload


def _wait_for_readiness(url: str, container_name: str) -> dict[str, object]:
    deadline = time.monotonic() + 90
    last_error = "web process did not answer"
    while time.monotonic() < deadline:
        try:
            payload = _get_json(f"{url}/readyz")
            if payload.get("docker_runtime_available") is True:
                return payload
            last_error = str(payload.get("runtime_error") or payload)
        except (OSError, urllib.error.URLError, json.JSONDecodeError) as exc:
            last_error = str(exc)
        time.sleep(0.5)
    logs = _docker("logs", container_name, timeout=30).stdout
    pytest.fail(f"web runtime did not become ready: {last_error}\n{logs}")


def _upload_xlsx(url: str, workbook_path: Path) -> dict[str, object]:
    boundary = f"xlsliberator-{uuid.uuid4().hex}"
    body = (
        (
            f"--{boundary}\r\n"
            'Content-Disposition: form-data; name="file"; filename="smoke.xlsx"\r\n'
            "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n"
            "\r\n"
        ).encode()
        + workbook_path.read_bytes()
        + f"\r\n--{boundary}--\r\n".encode()
    )
    request = urllib.request.Request(  # noqa: S310
        f"{url}/api/jobs",
        data=body,
        headers={
            "Accept": "application/json",
            "Content-Type": f"multipart/form-data; boundary={boundary}",
        },
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:  # noqa: S310
        payload = json.loads(response.read())
    assert isinstance(payload, dict)
    return payload


@pytest.mark.integration
@pytest.mark.docker
def test_web_container_converts_through_disposable_office_runtime(tmp_path: Path) -> None:
    if os.getenv("DOCKER_TESTS") != "1":
        pytest.skip("Set DOCKER_TESTS=1 to run Docker smoke tests")
    if shutil.which("docker") is None:
        pytest.skip("Docker is not installed")

    data_dir = tmp_path / "web-data"
    runtime_dir = tmp_path / "runtime-tmp"
    data_dir.mkdir(mode=0o777)
    runtime_dir.mkdir(mode=0o777)
    data_dir.chmod(0o777)
    runtime_dir.chmod(0o777)
    workbook_path = tmp_path / "smoke.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "Docker web smoke"
    workbook.active["A2"] = "=1+1"
    workbook.save(workbook_path)
    workbook.close()

    _docker(
        "build",
        "--file",
        "docker/office/libreoffice/Dockerfile",
        "--tag",
        "xlsliberator-libreoffice:26.2.4.2",
        ".",
    )
    _docker("build", "--tag", "xlsliberator-web:test", ".")

    container_name = f"xlsliberator-web-smoke-{uuid.uuid4().hex[:12]}"
    orchestrator_container = socket.gethostname()
    socket_gid = os.stat("/var/run/docker.sock").st_gid
    _docker(
        "run",
        "--detach",
        "--name",
        container_name,
        "--network",
        f"container:{orchestrator_container}",
        "--group-add",
        str(socket_gid),
        "--volume",
        "/var/run/docker.sock:/var/run/docker.sock",
        "--volume",
        f"{data_dir}:/data",
        "--volume",
        f"{runtime_dir}:/runtime-tmp",
        "--env",
        "XLSLIBERATOR_DATA_DIR=/data",
        "--env",
        "XLSLIBERATOR_RUNTIME_TEMP_ROOT=/runtime-tmp",
        "--env",
        f"XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT={runtime_dir}",
        "--env",
        "XLSLIBERATOR_WORKSPACE_ROOTS=/data",
        "--env",
        "XLSLIBERATOR_EMBED_MACROS=0",
        "--env",
        "XLSLIBERATOR_USE_AGENT=0",
        "xlsliberator-web:test",
    )
    try:
        base_url = "http://127.0.0.1:8080"
        readiness = _wait_for_readiness(base_url, container_name)
        assert readiness["version"] == "26.2.4.2"
        assert str(readiness["image_id"]).startswith("sha256:")

        queued = _upload_xlsx(base_url, workbook_path)
        job_id = str(queued["id"])
        deadline = time.monotonic() + 120
        status: dict[str, object] = queued
        while time.monotonic() < deadline:
            status = _get_json(f"{base_url}/api/jobs/{job_id}")
            if status.get("status") in {"completed", "failed"}:
                break
            time.sleep(0.5)
        assert status.get("status") == "completed", status
        output = data_dir / "jobs" / job_id / "output.ods"
        assert output.is_file()
        assert output.read_bytes().startswith(b"PK")
    finally:
        subprocess.run(
            ["docker", "rm", "--force", container_name],
            check=False,
            capture_output=True,
            text=True,
            timeout=30,
        )
