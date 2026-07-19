"""Docker smoke for the in-repository Open-SWE, MCP, and web services."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import time
import uuid
from pathlib import Path
from typing import Any

import openpyxl
import pytest

ROOT = Path(__file__).resolve().parents[2]
OVERRIDE = ROOT / "tests" / "integration" / "docker-compose.open-swe-smoke.yml"


def _compose(
    project: str,
    *arguments: str,
    timeout: int = 1200,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            "docker",
            "compose",
            "--project-name",
            project,
            "--file",
            str(ROOT / "docker-compose.yml"),
            "--file",
            str(OVERRIDE),
            *arguments,
        ],
        cwd=ROOT,
        check=check,
        capture_output=True,
        text=True,
        timeout=timeout,
    )


def _exec_python(project: str, service: str, program: str) -> str:
    result = subprocess.run(
        [
            "docker",
            "compose",
            "--project-name",
            project,
            "--file",
            str(ROOT / "docker-compose.yml"),
            "--file",
            str(OVERRIDE),
            "exec",
            "-T",
            service,
            "python",
            "-",
        ],
        cwd=ROOT,
        input=program,
        check=True,
        capture_output=True,
        text=True,
        timeout=60,
    )
    return result.stdout.strip()


def _wait_for_web(project: str) -> dict[str, Any]:
    program = """
import json
import urllib.request
print(json.dumps(json.load(urllib.request.urlopen("http://127.0.0.1:8080/readyz"))))
"""
    deadline = time.monotonic() + 120
    last_error = "web did not answer"
    while time.monotonic() < deadline:
        try:
            payload = json.loads(_exec_python(project, "xlsliberator-web", program))
            if payload.get("open_swe_reachable") is True:
                return payload
            last_error = str(payload)
        except (subprocess.CalledProcessError, json.JSONDecodeError) as exc:
            last_error = str(exc)
        time.sleep(0.5)
    logs = _compose(project, "logs", "--no-color", "--tail=200", check=False).stdout
    pytest.fail(f"embedded Open-SWE stack did not become ready: {last_error}\n{logs}")


@pytest.mark.integration
@pytest.mark.docker
def test_web_container_uses_embedded_open_swe_without_paid_default(tmp_path: Path) -> None:
    if os.getenv("DOCKER_TESTS") != "1":
        pytest.skip("Set DOCKER_TESTS=1 to run Docker smoke tests")
    if shutil.which("docker") is None:
        pytest.skip("Docker is not installed")

    project = f"xlsliberator-smoke-{uuid.uuid4().hex[:10]}"
    source = tmp_path / "smoke.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active["A1"] = "Embedded Open-SWE smoke"
    workbook.active["A2"] = "=1+1"
    workbook.save(source)
    workbook.close()

    environment = os.environ.copy()
    environment["XLSLIBERATOR_OPEN_SWE_MODEL"] = ""
    environment["XLSLIBERATOR_GITHUB_MODELS_ENABLED"] = "0"
    try:
        try:
            subprocess.run(
                [
                    "docker",
                    "compose",
                    "--project-name",
                    project,
                    "--file",
                    str(ROOT / "docker-compose.yml"),
                    "--file",
                    str(OVERRIDE),
                    "up",
                    "--detach",
                    "--build",
                    "xlsliberator-web",
                ],
                cwd=ROOT,
                env=environment,
                check=True,
                capture_output=True,
                text=True,
                timeout=1200,
            )
        except subprocess.CalledProcessError as exc:
            status = _compose(project, "ps", "--all", check=False)
            logs = _compose(project, "logs", "--no-color", "--tail=200", check=False)
            pytest.fail(
                "embedded Open-SWE Compose startup failed:\n"
                f"stdout:\n{exc.stdout}\n"
                f"stderr:\n{exc.stderr}\n"
                f"compose ps:\n{status.stdout}\n{status.stderr}\n"
                f"service logs:\n{logs.stdout}\n{logs.stderr}"
            )
        readiness = _wait_for_web(project)
        assert readiness["open_swe_configured"] is True
        assert readiness["target_libreoffice_version"] == "26.2.4.2"

        open_swe_health = json.loads(
            _exec_python(
                project,
                "xlsliberator-web",
                """
import json
import urllib.request
print(json.dumps(json.load(urllib.request.urlopen("http://xlsliberator-open-swe:2024/health"))))
""",
            )
        )
        assert open_swe_health["runtime"] == "open-swe"
        assert open_swe_health["model_configured"] is False
        assert open_swe_health["github_models_enabled"] is False

        graph = json.loads(
            _exec_python(
                project,
                "xlsliberator-open-swe",
                """
import asyncio
import json
import shutil
import uuid
from xlsliberator.open_swe_agent.graph import get_agent
from xlsliberator.open_swe_agent.state import thread_root, write_state

thread_id = str(uuid.uuid4())
write_state(thread_id, {"thread_id": thread_id, "events": []})
try:
    graph = asyncio.run(get_agent({"configurable": {"thread_id": thread_id}}))
    print(json.dumps({"type": type(graph).__name__, "nodes": sorted(graph.get_graph().nodes)}))
finally:
    shutil.rmtree(thread_root(thread_id))
""",
            )
        )
        assert graph["type"] == "CompiledStateGraph"
        assert {"model", "tools"}.issubset(graph["nodes"])

        tool_smoke = json.loads(
            _exec_python(
                project,
                "xlsliberator-open-swe",
                """
import asyncio
import json
import shutil
import uuid
from openpyxl import Workbook
from xlsliberator.open_swe_agent.state import thread_root, write_state
from xlsliberator.open_swe_agent.tools import workbook_tools

async def main():
    thread_id = str(uuid.uuid4())
    root = thread_root(thread_id)
    source = root / "source" / "tool-smoke.xlsx"
    source.parent.mkdir(parents=True)
    workbook = Workbook()
    workbook.active["A1"] = "Open-SWE MCP tool smoke"
    workbook.active["A2"] = "=1+1"
    workbook.save(source)
    workbook.close()
    write_state(thread_id, {"thread_id": thread_id, "events": []})
    try:
        tools = {tool.name: tool for tool in workbook_tools(thread_id)}
        result = await tools["inspect_source_workbook"].ainvoke(
            {"source_path": "/workspace/source/tool-smoke.xlsx"}
        )
        print(json.dumps({
            "operation_status": result.get("operation_status"),
            "evidence_written": (root / "evidence" / "source-inventory.json").is_file(),
        }))
    finally:
        shutil.rmtree(root, ignore_errors=True)

asyncio.run(main())
""",
            )
        )
        assert tool_smoke == {
            "operation_status": "passed",
            "evidence_written": True,
        }

        web_container = _compose(project, "ps", "--quiet", "xlsliberator-web").stdout.strip()
        subprocess.run(
            ["docker", "cp", str(source), f"{web_container}:/data/open-swe-smoke.xlsx"],
            check=True,
            capture_output=True,
            text=True,
            timeout=30,
        )
        queued = json.loads(
            _exec_python(
                project,
                "xlsliberator-web",
                """
import json
import uuid
import urllib.request
from pathlib import Path

boundary = "xlsliberator-" + uuid.uuid4().hex
source = Path("/data/open-swe-smoke.xlsx")
body = (
    (
        f"--{boundary}\\r\\n"
        'Content-Disposition: form-data; name="file"; filename="smoke.xlsx"\\r\\n'
        "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\\r\\n"
        "\\r\\n"
    ).encode()
    + source.read_bytes()
    + f"\\r\\n--{boundary}--\\r\\n".encode()
)
request = urllib.request.Request(
    "http://127.0.0.1:8080/api/jobs",
    data=body,
    headers={"Accept": "application/json", "Content-Type": f"multipart/form-data; boundary={boundary}"},
    method="POST",
)
print(urllib.request.urlopen(request).read().decode())
""",
            )
        )
        job_id = str(queued["id"])
        deadline = time.monotonic() + 30
        job: dict[str, Any] = queued
        while time.monotonic() < deadline:
            job = json.loads(
                _exec_python(
                    project,
                    "xlsliberator-web",
                    f"""
import urllib.request
print(urllib.request.urlopen("http://127.0.0.1:8080/api/jobs/{job_id}").read().decode())
""",
                )
            )
            if job.get("status") == "failed":
                break
            time.sleep(0.2)
        assert job["status"] == "failed"
        assert job["error"] == "Open-SWE migration service returned HTTP 503"
        assert job["thread_id"] is None

        open_swe_container = _compose(
            project,
            "ps",
            "--quiet",
            "xlsliberator-open-swe",
        ).stdout.strip()
        inspection = json.loads(
            subprocess.run(
                ["docker", "inspect", open_swe_container],
                check=True,
                capture_output=True,
                text=True,
                timeout=30,
            ).stdout
        )[0]
        assert all(mount["Destination"] != "/var/run/docker.sock" for mount in inspection["Mounts"])
    finally:
        _compose(project, "down", "--volumes", "--remove-orphans", timeout=120, check=False)
