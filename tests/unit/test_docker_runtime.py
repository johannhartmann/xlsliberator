"""Unit tests for immutable, disposable LibreOffice Docker jobs."""

from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.docker_runtime import (
    BASE_IMAGE_DIGEST,
    DockerRuntimeTimeout,
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
    MalformedWorkerResponse,
)


def _image_inspect(image_id: str = "sha256:fixed") -> str:
    return json.dumps(
        {
            "Id": image_id,
            "Architecture": "arm64",
            "Config": {
                "Labels": {
                    "org.xlsliberator.libreoffice.version": "26.2.4.2",
                    "org.xlsliberator.runtime.variant": "stock",
                }
            },
        }
    )


def _valid_probe() -> dict[str, Any]:
    return {
        "base_image_digest": BASE_IMAGE_DIGEST,
        "runtime_variant": "stock",
        "office_program_prefix": "/opt/libreoffice26.2/program/",
        "architecture": "aarch64",
        "uno_module": "/opt/libreoffice26.2/program/uno.py",
        "pyuno_native_module": "/opt/libreoffice26.2/program/pyuno.so",
        "python_executable": "/opt/libreoffice26.2/program/python.bin",
        "python_version": "3.12.13",
        "office_sha256": "a" * 64,
        "uno_module_sha256": "b" * 64,
        "pyuno_native_sha256": "c" * 64,
        "worker_wrapper_sha256": "d" * 64,
        "installed_package_manifest": [
            {
                "name": "libobasis26.2-pyuno",
                "version": "26.2.4.2-2",
                "architecture": "arm64",
            }
        ],
    }


def test_container_command_is_immutable_disposable_and_sandboxed(tmp_path: Path) -> None:
    command = LibreOfficeDockerRuntime()._container_command(  # noqa: SLF001
        "sha256:fixed", tmp_path, "xlsliberator-lo-job"
    )

    assert command[:3] == ["docker", "run", "--rm"]
    assert "--interactive" in command
    assert command[command.index("--name") + 1] == "xlsliberator-lo-job"
    assert command[command.index("--network") + 1] == "none"
    assert "--read-only" in command
    assert command[command.index("--cap-drop") + 1] == "ALL"
    assert command[command.index("--pids-limit") + 1] == "256"
    assert command[command.index("--memory") + 1] == "2g"
    assert command[command.index("--user") + 1] == "10001:10001"
    assert "sha256:fixed" in command
    assert "soffice" not in command
    assert "/var/run/docker.sock" not in " ".join(command)


def test_nested_orchestrator_translates_only_runtime_temp_paths(
    tmp_path: Path, monkeypatch: Any
) -> None:
    local_root = tmp_path / "container-runtime"
    job_dir = local_root / "job-1" / "job"
    input_dir = local_root / "job-1" / "input"
    job_dir.mkdir(parents=True)
    input_dir.mkdir()
    monkeypatch.setenv("XLSLIBERATOR_RUNTIME_TEMP_ROOT", str(local_root))
    monkeypatch.setenv(
        "XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT", "/srv/xlsliberator/runtime-tmp"
    )

    command = LibreOfficeDockerRuntime()._container_command(  # noqa: SLF001
        "sha256:fixed", job_dir, "xlsliberator-lo-job", input_dir=input_dir
    )

    mounts = [command[index + 1] for index, value in enumerate(command) if value == "--mount"]
    assert "type=bind,src=/srv/xlsliberator/runtime-tmp/job-1/input,dst=/input,readonly" in mounts
    assert "type=bind,src=/srv/xlsliberator/runtime-tmp/job-1/job,dst=/job" in mounts
    assert "/var/run/docker.sock" not in " ".join(command)


def test_nested_orchestrator_rejects_path_outside_runtime_root(
    tmp_path: Path, monkeypatch: Any
) -> None:
    runtime_root = tmp_path / "runtime"
    runtime_root.mkdir()
    outside = tmp_path / "outside"
    outside.mkdir()
    monkeypatch.setenv("XLSLIBERATOR_RUNTIME_TEMP_ROOT", str(runtime_root))
    monkeypatch.setenv("XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT", "/srv/runtime")

    with pytest.raises(DockerRuntimeUnavailable, match="escaped"):
        LibreOfficeDockerRuntime._docker_mount_source(outside)  # noqa: SLF001


def test_missing_docker_fails_without_host_fallback(monkeypatch: Any) -> None:
    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: None)

    def forbidden(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("no process may be started")

    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", forbidden)
    with pytest.raises(DockerRuntimeUnavailable, match="host fallback is disabled"):
        LibreOfficeDockerRuntime().resolve_identity()


def test_request_runs_resolved_image_id_and_records_job_identity(monkeypatch: Any) -> None:
    calls: list[list[str]] = []

    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        if command[1:3] == ["image", "inspect"] and command[-1] == "{{json .}}":
            stdout = _image_inspect()
        elif command[1] == "run":
            stdout = json.dumps({"success": True, "op": "ping", "data": {"uno_importable": True}})
        else:
            stdout = "sha256:fixed\n"
        return subprocess.CompletedProcess(command, 0, stdout=stdout, stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    response = LibreOfficeDockerRuntime().request(
        {
            "op": "ping",
            "environment": {
                "declared_capabilities": ["macro_execution"],
                "granted_capabilities": ["macro_execution"],
            },
        }
    )

    run_command = next(command for command in calls if command[1] == "run")
    assert "sha256:fixed" in run_command
    assert response["data"]["container_name"].startswith("xlsliberator-lo-")
    assert response["data"]["job_id"]
    assert response["data"]["resource_policy"]["network"] == "none"
    assert response["data"]["sandbox_job"]["image_digest"] == "sha256:fixed"
    assert response["data"]["granted_capabilities"] == ["macro_execution"]


def test_request_stages_working_copy_and_attachment_outputs(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    source = tmp_path / "source.ods"
    output = tmp_path / "working.ods"
    attachment = tmp_path / "export.pdf"
    source.write_bytes(b"source")

    def fake_run(
        command: list[str],
        **kwargs: Any,
    ) -> subprocess.CompletedProcess[str]:
        if command[1:3] == ["image", "inspect"] and command[-1] == "{{json .}}":
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=_image_inspect(),
                stderr="",
            )
        if command[1] == "run":
            payload = json.loads(str(kwargs["input"]))
            mounts = [
                command[index + 1] for index, value in enumerate(command) if value == "--mount"
            ]
            job_mount = next(item for item in mounts if item.endswith(",dst=/job"))
            job_root = Path(job_mount.split(",src=", 1)[1].split(",dst=", 1)[0])
            (job_root / Path(payload["output_path"]).name).write_bytes(b"working")
            (job_root / Path(payload["attachment_output_path"]).name).write_bytes(b"%PDF")
            return subprocess.CompletedProcess(
                command,
                0,
                stdout=json.dumps({"success": True, "op": "run_scenario", "data": {}}),
                stderr="",
            )
        return subprocess.CompletedProcess(
            command,
            0,
            stdout="sha256:fixed\n",
            stderr="",
        )

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    response = LibreOfficeDockerRuntime(workspace_roots=[tmp_path]).request(
        {
            "op": "run_scenario",
            "ods_path": str(source),
            "output_path": str(output),
            "attachment_output_path": str(attachment),
        }
    )

    assert response["success"] is True
    assert output.read_bytes() == b"working"
    assert attachment.read_bytes() == b"%PDF"


def test_image_tag_drift_invalidates_evidence(monkeypatch: Any) -> None:
    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        if command[1:3] == ["image", "inspect"] and command[-1] == "{{json .}}":
            stdout = _image_inspect()
        elif command[1] == "run":
            stdout = json.dumps({"success": True, "op": "ping", "data": {}})
        else:
            stdout = "sha256:moved\n"
        return subprocess.CompletedProcess(command, 0, stdout=stdout, stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    with pytest.raises(DockerRuntimeUnavailable, match="tag drifted"):
        LibreOfficeDockerRuntime().request({"op": "ping"})


def test_worker_error_json_is_preserved_when_container_exits_nonzero(monkeypatch: Any) -> None:
    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        if command[1:3] == ["image", "inspect"] and command[-1] == "{{json .}}":
            stdout = _image_inspect()
            return subprocess.CompletedProcess(command, 0, stdout=stdout, stderr="")
        if command[1] == "run":
            stdout = json.dumps(
                {
                    "success": False,
                    "op": "runtime_probe",
                    "data": {},
                    "error": {"type": "RuntimeError", "message": "probe failed"},
                }
            )
            return subprocess.CompletedProcess(command, 1, stdout=stdout, stderr="")
        return subprocess.CompletedProcess(command, 0, stdout="sha256:fixed\n", stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    response = LibreOfficeDockerRuntime().request({"op": "runtime_probe"})

    assert response["success"] is False
    assert response["error"]["message"] == "probe failed"


def test_wrong_version_label_blocks_runtime(monkeypatch: Any) -> None:
    data = json.loads(_image_inspect())
    data["Config"]["Labels"]["org.xlsliberator.libreoffice.version"] = "25.8.7.2"

    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0, stdout=json.dumps(data), stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    with pytest.raises(DockerRuntimeUnavailable, match="version mismatch"):
        LibreOfficeDockerRuntime().resolve_identity(probe=False)


def test_patched_image_cannot_masquerade_as_stock(monkeypatch: Any) -> None:
    data = json.loads(_image_inspect())
    data["Config"]["Labels"]["org.xlsliberator.runtime.variant"] = "patched-v1"

    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0, stdout=json.dumps(data), stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    with pytest.raises(DockerRuntimeUnavailable, match="variant mismatch"):
        LibreOfficeDockerRuntime().resolve_identity(probe=False)


def test_patched_variant_must_be_selected_explicitly(monkeypatch: Any) -> None:
    data = json.loads(_image_inspect())
    data["Config"]["Labels"]["org.xlsliberator.runtime.variant"] = "patched-v1"

    def fake_run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        return subprocess.CompletedProcess(command, 0, stdout=json.dumps(data), stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")
    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", fake_run)

    identity = LibreOfficeDockerRuntime(
        image="xlsliberator-libreoffice:patched-v1", expected_variant="patched-v1"
    ).resolve_identity(probe=False)
    assert identity.runtime_variant == "patched-v1"


def test_stock_source_probe_requires_explicit_unpatched_identity() -> None:
    probe = _valid_probe()
    probe.update(
        {
            "runtime_variant": "stock-source",
            "office_program_prefix": "/opt/libreoffice/program/",
            "uno_module": "/opt/libreoffice/program/uno.py",
            "pyuno_native_module": "/opt/libreoffice/program/pyuno.so",
            "python_executable": "/opt/libreoffice/program/python",
            "source_commit": "a" * 40,
            "patch_set_sha256": "none",
        }
    )

    LibreOfficeDockerRuntime._validate_probe_provenance(  # noqa: SLF001
        probe, "arm64", expected_variant="stock-source"
    )


def test_stock_source_probe_rejects_patched_identity() -> None:
    probe = _valid_probe()
    probe.update(
        {
            "runtime_variant": "stock-source",
            "office_program_prefix": "/opt/libreoffice/program/",
            "uno_module": "/opt/libreoffice/program/uno.py",
            "pyuno_native_module": "/opt/libreoffice/program/pyuno.so",
            "python_executable": "/opt/libreoffice/program/python",
            "source_commit": "a" * 40,
            "patch_set_sha256": "b" * 64,
        }
    )

    with pytest.raises(DockerRuntimeUnavailable, match="patch-set identity"):
        LibreOfficeDockerRuntime._validate_probe_provenance(  # noqa: SLF001
            probe, "arm64", expected_variant="stock-source"
        )


@pytest.mark.parametrize(
    ("mutation", "message"),
    [
        ({"base_image_digest": "sha256:wrong"}, "base-image digest"),
        ({"pyuno_native_module": "/usr/lib/python3/pyuno.so"}, "PyUNO provenance"),
        ({"python_executable": "/usr/bin/python3"}, "bundled Python"),
        ({"python_version": "3.11.2"}, "Python version"),
        ({"installed_package_manifest": []}, "installed-package manifest"),
        ({"architecture": "x86_64"}, "architecture"),
    ],
)
def test_runtime_probe_provenance_mismatch_blocks_readiness(
    mutation: dict[str, Any], message: str
) -> None:
    probe = _valid_probe()
    probe.update(mutation)

    with pytest.raises(DockerRuntimeUnavailable, match=message):
        LibreOfficeDockerRuntime._validate_probe_provenance(probe, "arm64")  # noqa: SLF001


def test_runtime_probe_requires_matching_pyuno_package() -> None:
    probe = _valid_probe()
    probe["installed_package_manifest"][0]["version"] = "25.8.7.2-2"

    with pytest.raises(DockerRuntimeUnavailable, match="does not match LibreOffice"):
        LibreOfficeDockerRuntime._validate_probe_provenance(probe, "arm64")  # noqa: SLF001


def test_worker_timeout_retains_specific_failure_type(monkeypatch: Any) -> None:
    """Wall-time failures must not be hidden as generic unavailability."""
    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")

    def time_out(*_args: Any, **_kwargs: Any) -> Any:
        raise subprocess.TimeoutExpired(["docker", "run"], 7)

    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", time_out)

    with pytest.raises(DockerRuntimeTimeout, match="7s wall-time limit"):
        LibreOfficeDockerRuntime()._run_docker_cli(  # noqa: SLF001
            ["docker", "run"], timeout_seconds=7
        )


def test_worker_timeout_force_removes_named_container_process_tree(monkeypatch: Any) -> None:
    """A timed-out office job must be force-removed before returning."""
    calls: list[list[str]] = []
    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")

    def run(command: list[str], **_kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        if command[1] == "run":
            raise subprocess.TimeoutExpired(command, 3)
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", run)
    command = ["docker", "run", "--name", "xlsliberator-lo-timeout", "sha256:fixed"]

    with pytest.raises(DockerRuntimeTimeout):
        LibreOfficeDockerRuntime()._run_docker_cli(  # noqa: SLF001
            command,
            timeout_seconds=3,
        )

    assert ["docker", "rm", "--force", "xlsliberator-lo-timeout"] in calls


def test_runtime_process_boundary_rejects_every_non_docker_command(monkeypatch: Any) -> None:
    monkeypatch.setattr("xlsliberator.docker_runtime.shutil.which", lambda _name: "/bin/docker")

    def forbidden(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("a rejected host process must never start")

    monkeypatch.setattr("xlsliberator.docker_runtime.subprocess.run", forbidden)
    with pytest.raises(DockerRuntimeUnavailable, match="only the configured Docker CLI"):
        LibreOfficeDockerRuntime()._run_docker_cli(  # noqa: SLF001
            ["python3", "-c", "import uno"], timeout_seconds=1
        )


def test_runtime_rejects_non_docker_executable_configuration() -> None:
    with pytest.raises(DockerRuntimeUnavailable, match="Only the Docker CLI"):
        LibreOfficeDockerRuntime(docker_executable="python3").resolve_identity(probe=False)


def test_malformed_worker_response_retains_protocol_failure_type() -> None:
    """Malformed JSON is a protocol failure, not an empty successful result."""
    result = subprocess.CompletedProcess(
        ["docker", "run"], 0, stdout="not-json", stderr="container diagnostic"
    )

    with pytest.raises(MalformedWorkerResponse, match="malformed JSON"):
        LibreOfficeDockerRuntime._parse_response(result)  # noqa: SLF001
