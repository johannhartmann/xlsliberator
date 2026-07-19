"""Execution sandbox policy shared by office, oracle, macro, GUI, and agent jobs."""

from __future__ import annotations

import os
import stat
import subprocess
import time
from enum import StrEnum
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from xlsliberator.validation_models import GateExecutionStatus


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class SandboxUnavailable(RuntimeError):
    """A required execution sandbox is not usable."""


class WorkspaceAccessError(ValueError):
    """A requested host path escaped the configured workspace roots."""


class ExecutionKind(StrEnum):
    SOURCE_ORACLE = "source_oracle"
    LIBREOFFICE_TARGET = "libreoffice_target"
    MACRO = "macro"
    GUI = "gui"
    CODING_AGENT_BUILD = "coding_agent_build"
    CODING_AGENT_TEST = "coding_agent_test"


class SandboxBackendKind(StrEnum):
    DOCKER = "docker"
    REMOTE_WORKER = "remote_worker"
    MICROVM = "microvm"


class SandboxLimits(StrictModel):
    cpu_count: float = Field(default=2.0, gt=0)
    memory_bytes: int = Field(default=2 * 1024**3, ge=64 * 1024**2)
    shared_memory_bytes: int = Field(default=256 * 1024**2, ge=64 * 1024**2)
    process_count: int = Field(default=256, ge=1)
    file_size_bytes: int = Field(default=1024**3, ge=1024)
    wall_seconds: float = Field(default=120.0, gt=0)
    writable_bytes: int = Field(default=1024**3, ge=1024)


class SandboxPolicy(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    backend: SandboxBackendKind = SandboxBackendKind.DOCKER
    network: Literal["none"] = "none"
    read_only_root: Literal[True] = True
    inherited_credentials: Literal[False] = False
    disposable_user: Literal[True] = True
    disposable_home: Literal[True] = True
    restricted_devices: Literal[True] = True
    isolated_ipc: Literal[True] = True
    process_tree_cleanup: Literal[True] = True
    docker_socket: Literal[False] = False
    limits: SandboxLimits = Field(default_factory=SandboxLimits)

    def evidence(self) -> dict[str, object]:
        """Stable policy fields embedded in runtime evidence."""

        return self.model_dump(mode="json")


class SandboxMount(StrictModel):
    source: str
    destination: str
    mode: Literal["ro", "rw"]
    purpose: Literal["input", "job"]

    @model_validator(mode="after")
    def only_job_mount_is_writable(self) -> SandboxMount:
        if self.mode == "rw" and self.purpose != "job":
            raise ValueError("only the isolated job mount may be writable")
        if self.destination == "/var/run/docker.sock" or "docker.sock" in self.source:
            raise ValueError("the Docker socket cannot cross the sandbox boundary")
        return self


class SandboxJob(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    job_id: str
    kind: ExecutionKind
    image_reference: str
    image_digest: str
    mounts: list[SandboxMount]
    granted_capabilities: list[str] = Field(default_factory=list)
    policy: SandboxPolicy = Field(default_factory=SandboxPolicy)

    @model_validator(mode="after")
    def immutable_and_explicit(self) -> SandboxJob:
        if not self.image_digest.startswith("sha256:"):
            raise ValueError("sandbox jobs require a resolved immutable image digest")
        if sum(mount.purpose == "job" for mount in self.mounts) != 1:
            raise ValueError("sandbox jobs require exactly one isolated writable job mount")
        return self


class SandboxExecutionResult(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    status: GateExecutionStatus
    job: SandboxJob
    exit_code: int | None = None
    duration_seconds: float = Field(ge=0)
    stdout: str = ""
    stderr: str = ""
    error: str | None = None


class DockerCommandSandbox:
    """Execute a trusted command inside a pre-resolved disposable Docker job."""

    def __init__(
        self,
        *,
        docker_executable: str = "docker",
        workspace_roots: list[Path] | tuple[Path, ...] | None = None,
    ) -> None:
        self.docker_executable = docker_executable
        self.workspace = WorkspacePathPolicy(workspace_roots)

    def execute(self, job: SandboxJob, command: list[str]) -> SandboxExecutionResult:
        if job.policy.backend is not SandboxBackendKind.DOCKER:
            raise SandboxUnavailable("Docker command sandbox requires the Docker backend")
        if not command or any("\x00" in argument for argument in command):
            raise ValueError("sandbox command is empty or malformed")
        self._require_docker()
        mounts: list[str] = []
        for mount in job.mounts:
            source = Path(mount.source).resolve(strict=True)
            if not any(
                source == root or source.is_relative_to(root) for root in self.workspace.roots
            ):
                raise WorkspaceAccessError(f"sandbox mount escaped workspace roots: {source}")
            suffix = ",readonly" if mount.mode == "ro" else ""
            mounts.extend(
                [
                    "--mount",
                    f"type=bind,src={source},dst={mount.destination}{suffix}",
                ]
            )
        container_name = f"xlsliberator-sandbox-{job.job_id}"
        argv = [
            self.docker_executable,
            "run",
            "--rm",
            "--name",
            container_name,
            *docker_sandbox_arguments(job.policy),
            *mounts,
            job.image_digest,
            *command,
        ]
        started = time.monotonic()
        try:
            result = subprocess.run(
                argv,
                capture_output=True,
                text=True,
                check=False,
                timeout=job.policy.limits.wall_seconds,
                env={"PATH": os.environ.get("PATH", "")},
            )
        except subprocess.TimeoutExpired:
            self._remove(container_name)
            return SandboxExecutionResult(
                status=GateExecutionStatus.UNAVAILABLE,
                job=job,
                duration_seconds=time.monotonic() - started,
                error="sandbox wall-time limit exceeded",
            )
        return SandboxExecutionResult(
            status=(
                GateExecutionStatus.PASSED if result.returncode == 0 else GateExecutionStatus.FAILED
            ),
            job=job,
            exit_code=result.returncode,
            duration_seconds=time.monotonic() - started,
            stdout=result.stdout[-64_000:],
            stderr=result.stderr[-64_000:],
            error=None if result.returncode == 0 else "sandbox command failed",
        )

    def _require_docker(self) -> None:
        from shutil import which

        if which(self.docker_executable) is None:
            raise SandboxUnavailable("Docker sandbox is unavailable; host fallback is forbidden")

    def _remove(self, container_name: str) -> None:
        try:
            subprocess.run(
                [self.docker_executable, "rm", "--force", container_name],
                capture_output=True,
                text=True,
                check=False,
                timeout=15,
                env={"PATH": os.environ.get("PATH", "")},
            )
        except (OSError, subprocess.SubprocessError):
            return


class WorkspacePathPolicy:
    """Resolve host paths beneath explicit roots without symlink traversal."""

    def __init__(self, roots: list[Path] | tuple[Path, ...] | None = None) -> None:
        configured = list(roots or self._environment_roots())
        if not configured:
            configured = [Path.cwd()]
        resolved: list[Path] = []
        for root in configured:
            candidate = root.expanduser().resolve(strict=True)
            if not candidate.is_dir():
                raise WorkspaceAccessError(f"workspace root is not a directory: {candidate}")
            resolved.append(candidate)
        self.roots = tuple(dict.fromkeys(resolved))

    @staticmethod
    def _environment_roots() -> list[Path]:
        raw = os.environ.get("XLSLIBERATOR_WORKSPACE_ROOTS", "")
        return [Path(item) for item in raw.split(os.pathsep) if item]

    def input_file(self, path: str | Path) -> Path:
        candidate = Path(path).expanduser()
        self._reject_symlink_components(candidate)
        resolved = candidate.resolve(strict=True)
        if not resolved.is_file() or not self._beneath_root(resolved):
            raise WorkspaceAccessError(f"input file is outside configured workspace roots: {path}")
        mode = resolved.stat().st_mode
        if not stat.S_ISREG(mode):
            raise WorkspaceAccessError(f"input is not a regular file: {path}")
        return resolved

    def output_file(self, path: str | Path) -> Path:
        candidate = Path(path).expanduser()
        parent = candidate.parent.resolve(strict=True)
        self._reject_symlink_components(candidate.parent)
        resolved = parent / candidate.name
        if not self._beneath_root(resolved):
            raise WorkspaceAccessError(f"output file is outside configured workspace roots: {path}")
        if resolved.exists() and resolved.is_symlink():
            raise WorkspaceAccessError(f"output cannot replace a symlink: {path}")
        return resolved

    def _beneath_root(self, path: Path) -> bool:
        return any(path == root or path.is_relative_to(root) for root in self.roots)

    @staticmethod
    def _reject_symlink_components(path: Path) -> None:
        absolute = path.absolute()
        current = Path(absolute.anchor)
        for component in absolute.parts[1:]:
            current /= component
            if current.exists() and current.is_symlink():
                raise WorkspaceAccessError(f"symlink traversal is forbidden: {current}")


def docker_sandbox_arguments(policy: SandboxPolicy) -> list[str]:
    """Return the mandatory Docker flags for one untrusted execution job."""

    limits = policy.limits
    mebibytes = max(64, limits.memory_bytes // 1024**2)
    memory_limit = (
        f"{limits.memory_bytes // 1024**3}g"
        if limits.memory_bytes % 1024**3 == 0
        else f"{mebibytes}m"
    )
    file_blocks = max(1, limits.file_size_bytes // 1024)
    writable_mebibytes = max(1, limits.writable_bytes // 1024**2)
    shared_memory_mebibytes = max(64, limits.shared_memory_bytes // 1024**2)
    return [
        "--network",
        "none",
        "--read-only",
        "--cap-drop",
        "ALL",
        "--security-opt",
        "no-new-privileges",
        "--pids-limit",
        str(limits.process_count),
        "--ulimit",
        f"fsize={file_blocks}:{file_blocks}",
        "--memory",
        memory_limit,
        "--cpus",
        str(limits.cpu_count),
        "--ipc",
        "private",
        "--shm-size",
        f"{shared_memory_mebibytes}m",
        "--init",
        "--stop-timeout",
        "5",
        "--tmpfs",
        # This is a private tmpfs inside the disposable container, not a host temp directory.
        f"/tmp:rw,noexec,nosuid,nodev,size={writable_mebibytes}m,mode=1777",  # nosec B108
        "--tmpfs",
        # Xorg requires this socket directory to be root-owned and sticky. Mount it
        # beneath the private /tmp before dropping to the disposable runtime user.
        "/tmp/.X11-unix:rw,noexec,nosuid,nodev,size=1m,mode=1777,uid=0,gid=0",
        "--tmpfs",
        "/home/sandbox:rw,noexec,nosuid,nodev,size=16m,mode=700,uid=10001,gid=10001",
        "--env",
        "HOME=/home/sandbox",
        "--env",
        "TMPDIR=/tmp",
        "--env",
        "LANG=C.UTF-8",
        "--user",
        "10001:10001",
    ]
