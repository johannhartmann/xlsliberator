"""UNO-free sandbox, malicious-input, and workspace-boundary tests."""

from __future__ import annotations

import stat
import zipfile
from pathlib import Path

import pytest
from pydantic import ValidationError

from xlsliberator.docker_runtime import LibreOfficeDockerRuntime
from xlsliberator.execution_sandbox import (
    ExecutionKind,
    SandboxJob,
    SandboxMount,
    SandboxPolicy,
    WorkspaceAccessError,
    WorkspacePathPolicy,
    docker_sandbox_arguments,
)
from xlsliberator.workbook_security import (
    UnsafeWorkbookError,
    WorkbookInputLimits,
    delimit_workbook_text,
    validate_untrusted_workbook,
)


def test_default_docker_policy_has_every_required_isolation_control() -> None:
    policy = SandboxPolicy()
    arguments = docker_sandbox_arguments(policy)
    rendered = " ".join(arguments)

    assert arguments[arguments.index("--network") + 1] == "none"
    assert "--read-only" in arguments
    assert arguments[arguments.index("--cap-drop") + 1] == "ALL"
    assert "no-new-privileges" in rendered
    assert "--pids-limit" in arguments
    assert "--memory" in arguments
    assert "--cpus" in arguments
    assert arguments[arguments.index("--ulimit") + 1] == "fsize=1073741824:1073741824"
    assert arguments[arguments.index("--ipc") + 1] == "private"
    assert arguments[arguments.index("--shm-size") + 1] == "256m"
    assert "--init" in arguments
    assert "HOME=/home/sandbox" in arguments
    assert "--user" in arguments
    assert not any(name in rendered for name in ("OPENAI_API_KEY", "ANTHROPIC_API_KEY"))
    assert "docker.sock" not in rendered


def test_sandbox_job_requires_digest_and_rejects_docker_socket(tmp_path: Path) -> None:
    job_mount = SandboxMount(source=str(tmp_path), destination="/job", mode="rw", purpose="job")
    with pytest.raises(ValidationError, match="resolved immutable image digest"):
        SandboxJob(
            job_id="job",
            kind=ExecutionKind.MACRO,
            image_reference="mutable:latest",
            image_digest="mutable:latest",
            mounts=[job_mount],
        )
    with pytest.raises(ValidationError, match="Docker socket"):
        SandboxMount(
            source="/var/run/docker.sock",
            destination="/var/run/docker.sock",
            mode="ro",
            purpose="input",
        )


def test_runtime_uses_separate_readonly_input_and_writable_job_mounts(tmp_path: Path) -> None:
    runtime = LibreOfficeDockerRuntime(workspace_roots=[tmp_path])
    command = runtime._container_command(  # noqa: SLF001
        "sha256:" + "a" * 64,
        tmp_path / "job",
        "sandbox-job",
        input_dir=tmp_path / "input",
    )
    rendered = " ".join(command)

    assert f"src={tmp_path / 'input'},dst=/input,readonly" in rendered
    assert f"src={tmp_path / 'job'},dst=/job" in rendered
    assert "/var/run/docker.sock" not in rendered
    assert "soffice" not in rendered


def test_workspace_policy_rejects_traversal_symlinks_and_outside_roots(tmp_path: Path) -> None:
    root = tmp_path / "workspace"
    outside = tmp_path / "outside"
    root.mkdir()
    outside.mkdir()
    secret = outside / "secret.xlsx"
    secret.write_bytes(b"secret")
    link = root / "link.xlsx"
    link.symlink_to(secret)
    policy = WorkspacePathPolicy([root])

    with pytest.raises(WorkspaceAccessError):
        policy.input_file(root / ".." / "outside" / "secret.xlsx")
    with pytest.raises(WorkspaceAccessError, match="symlink"):
        policy.input_file(link)
    with pytest.raises(WorkspaceAccessError):
        policy.output_file(outside / "stolen.ods")


def test_package_path_traversal_is_rejected(tmp_path: Path) -> None:
    workbook = tmp_path / "traversal.xlsx"
    with zipfile.ZipFile(workbook, "w") as archive:
        archive.writestr("../escape.xml", "<x/>")

    with pytest.raises(UnsafeWorkbookError, match="unsafe workbook package path"):
        validate_untrusted_workbook(workbook)


def test_package_symlink_is_rejected(tmp_path: Path) -> None:
    workbook = tmp_path / "symlink.xlsx"
    with zipfile.ZipFile(workbook, "w") as archive:
        info = zipfile.ZipInfo("xl/evil-link")
        info.create_system = 3
        info.external_attr = (stat.S_IFLNK | 0o777) << 16
        archive.writestr(info, "/etc/passwd")

    with pytest.raises(UnsafeWorkbookError, match="symlink"):
        validate_untrusted_workbook(workbook)


def test_zip_bomb_and_oversized_formula_are_rejected(tmp_path: Path) -> None:
    bomb = tmp_path / "bomb.xlsx"
    with zipfile.ZipFile(bomb, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("xl/worksheets/sheet1.xml", b"0" * 200_000)
    with pytest.raises(UnsafeWorkbookError, match="compression-ratio"):
        validate_untrusted_workbook(
            bomb,
            WorkbookInputLimits(max_compression_ratio=2.0, max_part_bytes=1_000_000),
        )

    oversized = tmp_path / "formula.xlsx"
    with zipfile.ZipFile(oversized, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", b"<f>" + b"A" * 100 + b"</f>")
    with pytest.raises(UnsafeWorkbookError, match="formula exceeds"):
        validate_untrusted_workbook(oversized, WorkbookInputLimits(max_formula_characters=20))


def test_oversized_macro_is_rejected(tmp_path: Path) -> None:
    workbook = tmp_path / "macro.ods"
    with zipfile.ZipFile(workbook, "w") as archive:
        archive.writestr("Scripts/python/evil.py", "x" * 100)

    with pytest.raises(UnsafeWorkbookError, match="macro source exceeds"):
        validate_untrusted_workbook(workbook, WorkbookInputLimits(max_macro_characters=20))


def test_malicious_fixture_text_stays_data_and_cannot_change_policy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    secret = "unit-test-openai-secret"
    monkeypatch.setenv("OPENAI_API_KEY", secret)
    fixtures = Path(__file__).parents[1] / "fixtures" / "security"
    policy_before = SandboxPolicy()
    for path in fixtures.iterdir():
        envelope = delimit_workbook_text(path.read_text(encoding="utf-8"), source=path.name)
        assert envelope.startswith("<UNTRUSTED_WORKBOOK_DATA")
        assert envelope.endswith("</UNTRUSTED_WORKBOOK_DATA>")
    assert SandboxPolicy() == policy_before
    assert secret not in " ".join(docker_sandbox_arguments(policy_before))


def test_mcp_is_loopback_only_in_trusted_local_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from xlsliberator import mcp_server

    calls: list[dict[str, object]] = []
    monkeypatch.setattr(mcp_server.mcp, "run", lambda **kwargs: calls.append(kwargs))

    mcp_server.serve()
    assert calls[0]["host"] == "127.0.0.1"
    with pytest.raises(ValueError, match="loopback"):
        mcp_server.serve(host="0.0.0.0")
    monkeypatch.setenv("XLSLIBERATOR_MCP_TRUSTED_CONTAINER_PROXY", "1")
    mcp_server.serve(host="0.0.0.0")
    assert calls[-1]["host"] == "0.0.0.0"
    with pytest.raises(ValueError, match="trusted-local"):
        mcp_server.serve(trusted_local=False)
