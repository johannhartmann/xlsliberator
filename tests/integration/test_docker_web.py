"""End-to-end smoke test for the Docker web client and a fake Open-SWE API."""

from __future__ import annotations

import base64
import io
import json
import os
import shutil
import subprocess
import threading
import time
import urllib.error
import urllib.request
import uuid
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

import openpyxl
import pytest

_THREAD_ID = "aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa"
_ARTIFACTS = {
    "a" * 24: (
        "target.ods",
        "ods",
        "application/vnd.oasis.opendocument.spreadsheet",
    ),
    "b" * 24: ("bridge.py", "generated", "text/x-python"),
    "c" * 24: ("save-reopen.json", "evidence", "application/json"),
}


def _ods_bytes() -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
    return output.getvalue()


_ARTIFACT_CONTENT = {
    "a" * 24: _ods_bytes(),
    "b" * 24: b"def migrate():\n    return True\n",
    "c" * 24: b'{"status":"passed"}',
}


class _FakeOpenSWEHandler(BaseHTTPRequestHandler):
    created_payload: dict[str, Any] | None = None
    request_headers: dict[str, str] = {}

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path)
        if path.path == "/health":
            self._json({"status": "healthy"})
            return
        if path.path == f"/api/xlsliberator/migrations/{_THREAD_ID}/events":
            self._require_auth()
            self._json(
                {
                    "thread_id": _THREAD_ID,
                    "events": [
                        {
                            "index": index,
                            "stage": stage,
                            "message": message,
                            "status": "complete",
                        }
                        for index, (stage, message) in enumerate(
                            [
                                ("upload", "Workbook accepted"),
                                ("lead", "Migration lead started"),
                                ("plan", "Behavioral plan ready"),
                                ("specialists", "Specialist tasks complete"),
                                ("libreoffice", "LibreOffice scenarios passed"),
                                ("reviewer", "Independent review approved"),
                                ("final", "Final evidence ready"),
                            ]
                        )
                    ],
                    "next": 7,
                }
            )
            return
        if path.path == f"/api/xlsliberator/migrations/{_THREAD_ID}":
            self._require_auth()
            self._json(
                {
                    "thread_id": _THREAD_ID,
                    "run_id": "run-1",
                    "status": "complete",
                    "artifacts": [
                        {
                            "id": artifact_id,
                            "name": values[0],
                            "kind": values[1],
                            "media_type": values[2],
                            "size": len(_ARTIFACT_CONTENT[artifact_id]),
                        }
                        for artifact_id, values in _ARTIFACTS.items()
                    ],
                }
            )
            return
        prefix = f"/api/xlsliberator/migrations/{_THREAD_ID}/artifacts/"
        if path.path.startswith(prefix):
            self._require_auth()
            artifact_id = path.path.removeprefix(prefix)
            content = _ARTIFACT_CONTENT.get(artifact_id)
            if content is None:
                self.send_error(404)
                return
            self.send_response(200)
            self.send_header("Content-Length", str(len(content)))
            self.end_headers()
            self.wfile.write(content)
            return
        self.send_error(404)

    def do_POST(self) -> None:  # noqa: N802
        if self.path != "/api/xlsliberator/migrations":
            self.send_error(404)
            return
        self._require_auth()
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length))
        assert isinstance(payload, dict)
        type(self).created_payload = payload
        type(self).request_headers = dict(self.headers.items())
        self._json(
            {
                "thread_id": _THREAD_ID,
                "run_id": "run-1",
                "duplicate": False,
                "artifact_locations": {},
            },
            status=202,
        )

    def log_message(self, _format: str, *_args: object) -> None:
        return

    def _require_auth(self) -> None:
        if (
            self.headers.get("Authorization") != "Bearer smoke-secret"
            or self.headers.get("X-XLSLiberator-Owner") != "smoke-owner"
        ):
            self.send_error(401)
            raise ConnectionAbortedError("invalid fake-service authentication")

    def _json(self, payload: dict[str, Any], *, status: int = 200) -> None:
        content = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)


def _docker(*arguments: str, timeout: int = 600) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["docker", *arguments],
        check=True,
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def _get_json(url: str) -> dict[str, Any]:
    with urllib.request.urlopen(url, timeout=15) as response:  # noqa: S310
        payload = json.loads(response.read())
    assert isinstance(payload, dict)
    return payload


def _wait_for_readiness(url: str, container_name: str) -> dict[str, Any]:
    deadline = time.monotonic() + 90
    last_error = "web process did not answer"
    while time.monotonic() < deadline:
        try:
            payload = _get_json(f"{url}/readyz")
            if payload.get("open_swe_reachable") is True:
                return payload
            last_error = str(payload)
        except (OSError, urllib.error.URLError, json.JSONDecodeError) as exc:
            last_error = str(exc)
        time.sleep(0.5)
    logs = _docker("logs", container_name, timeout=30).stdout
    pytest.fail(f"web client did not become ready: {last_error}\n{logs}")


def _upload_xlsx(url: str, workbook_path: Path) -> dict[str, Any]:
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
def test_web_container_delegates_to_authenticated_open_swe(tmp_path: Path) -> None:
    if os.getenv("DOCKER_TESTS") != "1":
        pytest.skip("Set DOCKER_TESTS=1 to run Docker smoke tests")
    if shutil.which("docker") is None:
        pytest.skip("Docker is not installed")

    data_dir = tmp_path / "web-data"
    data_dir.mkdir(mode=0o777)
    data_dir.chmod(0o777)
    workbook_path = tmp_path / "smoke.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "Docker web smoke"
    workbook.active["A2"] = "=1+1"
    workbook.save(workbook_path)
    workbook.close()

    fake_server = ThreadingHTTPServer(("0.0.0.0", 0), _FakeOpenSWEHandler)
    fake_thread = threading.Thread(target=fake_server.serve_forever, daemon=True)
    fake_thread.start()
    fake_port = fake_server.server_address[1]

    _docker("build", "--tag", "xlsliberator-web:test", ".")
    container_name = f"xlsliberator-web-smoke-{uuid.uuid4().hex[:12]}"
    orchestrator_container = os.uname().nodename
    _docker(
        "run",
        "--detach",
        "--name",
        container_name,
        "--network",
        f"container:{orchestrator_container}",
        "--volume",
        f"{data_dir}:/data",
        "--env",
        "XLSLIBERATOR_DATA_DIR=/data",
        "--env",
        f"XLSLIBERATOR_OPEN_SWE_URL=http://127.0.0.1:{fake_port}",
        "--env",
        "XLSLIBERATOR_OPEN_SWE_TOKEN=smoke-secret",
        "--env",
        "XLSLIBERATOR_OPEN_SWE_OWNER_ID=smoke-owner",
        "--env",
        "XLSLIBERATOR_OPEN_SWE_POLL_SECONDS=0.1",
        "--env",
        "XLSLIBERATOR_OPEN_SWE_REQUEST_TIMEOUT_SECONDS=10",
        "xlsliberator-web:test",
    )
    try:
        base_url = "http://127.0.0.1:8080"
        readiness = _wait_for_readiness(base_url, container_name)
        assert readiness["open_swe_configured"] is True
        assert readiness["target_libreoffice_version"] == "26.2.4.2"

        mounts = json.loads(_docker("inspect", container_name).stdout)[0]["Mounts"]
        assert [mount["Destination"] for mount in mounts] == ["/data"]

        queued = _upload_xlsx(base_url, workbook_path)
        job_id = str(queued["id"])
        deadline = time.monotonic() + 60
        job: dict[str, Any] = queued
        while time.monotonic() < deadline:
            job = _get_json(f"{base_url}/api/jobs/{job_id}")
            if job.get("status") in {"completed", "failed"}:
                break
            time.sleep(0.2)

        assert job.get("status") == "completed", job
        assert job["thread_id"] == _THREAD_ID
        assert {artifact["kind"] for artifact in job["artifacts"]} == {
            "ods",
            "generated",
            "evidence",
        }
        assert all(
            stage in {event["step"] for event in job["events"]}
            for stage in ("lead", "specialists", "libreoffice", "reviewer", "final")
        )
        assert str(tmp_path) not in json.dumps(job)

        output = data_dir / "jobs" / job_id / "output.ods"
        source = data_dir / "jobs" / job_id / "input.xlsx"
        assert zipfile.is_zipfile(output)
        assert not source.exists()

        payload = _FakeOpenSWEHandler.created_payload
        assert payload is not None
        assert payload["owner_id"] == "smoke-owner"
        assert payload["target_libreoffice_version"] == "26.2.4.2"
        assert base64.b64decode(payload["artifact"]["artifact_base64"]).startswith(b"PK")
        assert _FakeOpenSWEHandler.request_headers["Authorization"] == "Bearer smoke-secret"
    finally:
        subprocess.run(
            ["docker", "rm", "--force", container_name],
            check=False,
            capture_output=True,
            text=True,
            timeout=30,
        )
        fake_server.shutdown()
        fake_server.server_close()
        fake_thread.join(timeout=5)
