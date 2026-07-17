"""Docker-only LibreOffice target scenario execution."""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Protocol
from uuid import uuid4

from xlsliberator.docker_runtime import (
    BASE_IMAGE_DIGEST,
    DockerRuntimeIdentity,
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    StepResult,
)
from xlsliberator.validation_models import GateExecutionStatus


class CalcTargetRunner(Protocol):
    """Common interface for Calc-compatible target runtime jobs."""

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> RuntimeTrace:
        """Run a scenario against a target without mutating ``workbook``."""


class DockerScenarioRuntime(Protocol):
    """Narrow orchestration boundary used by the target runner."""

    def resolve_identity(self, *, probe: bool = True) -> DockerRuntimeIdentity:
        """Resolve and probe the immutable runtime identity."""

    def request(self, payload: dict[str, Any], *, _identity: str | None = None) -> dict[str, Any]:
        """Run one isolated worker request."""


class LibreOfficeScenarioRunner:
    """Execute one scenario in one disposable pinned LibreOffice container."""

    def __init__(
        self,
        runtime: DockerScenarioRuntime | None = None,
        *,
        timeout_seconds: int = 120,
    ) -> None:
        self.runtime = runtime or LibreOfficeDockerRuntime(timeout_seconds=timeout_seconds)
        self.timeout_seconds = timeout_seconds

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> RuntimeTrace:
        workbook = workbook.resolve()
        before_hash = _hash_file(workbook)
        started_at = datetime.now(UTC)
        try:
            identity = self.runtime.resolve_identity(probe=True)
        except DockerRuntimeUnavailable as exc:
            return self._unavailable_trace(
                workbook,
                environment,
                scenario,
                before_hash,
                started_at,
                exc,
            )

        try:
            response = self.runtime.request(
                {
                    "op": "run_scenario",
                    "ods_path": str(workbook),
                    "scenario": scenario.model_dump(mode="json"),
                    "environment": environment.model_dump(mode="json"),
                    "timeout_seconds": self.timeout_seconds,
                },
                _identity=identity.image_id,
            )
        except DockerRuntimeUnavailable as exc:
            return self._unavailable_trace(
                workbook,
                environment,
                scenario,
                before_hash,
                started_at,
                exc,
                identity=identity,
            )

        after_hash = _hash_file(workbook)
        data = dict(response.get("data") or {})
        runtime_identity = _runtime_identity(identity, data)
        if not response.get("success"):
            error = dict(response.get("error") or {})
            return RuntimeTrace(
                trace_id=f"libreoffice-{uuid4().hex}",
                scenario_id=scenario.id,
                runtime_role="target",
                runtime_identity=runtime_identity,
                environment=environment,
                status=GateExecutionStatus.FAILED,
                started_at=started_at,
                ended_at=datetime.now(UTC),
                workbook_hash_before=before_hash,
                workbook_hash_after=after_hash,
                logs=_runtime_logs(data),
                error=error or {"type": "WorkerError", "message": "worker request failed"},
            )

        try:
            steps = [StepResult.model_validate(item) for item in data.get("steps") or []]
            status = GateExecutionStatus(str(data.get("status") or "failed"))
        except (TypeError, ValueError) as exc:
            return RuntimeTrace(
                trace_id=f"libreoffice-{uuid4().hex}",
                scenario_id=scenario.id,
                runtime_role="target",
                runtime_identity=runtime_identity,
                environment=environment,
                status=GateExecutionStatus.FAILED,
                started_at=started_at,
                ended_at=datetime.now(UTC),
                workbook_hash_before=before_hash,
                workbook_hash_after=after_hash,
                logs=_runtime_logs(data),
                error={"type": "TraceValidationError", "message": str(exc)},
            )

        source_mutated = before_hash != after_hash or bool(data.get("source_mutated"))
        if source_mutated:
            status = GateExecutionStatus.FAILED
        return RuntimeTrace(
            trace_id=f"libreoffice-{uuid4().hex}",
            scenario_id=scenario.id,
            runtime_role="target",
            runtime_identity=runtime_identity,
            environment=environment,
            status=status,
            started_at=datetime.fromisoformat(str(data["started_at"])),
            ended_at=datetime.fromisoformat(str(data["ended_at"])),
            workbook_hash_before=before_hash,
            workbook_hash_after=after_hash,
            steps=steps,
            logs=_runtime_logs(data),
            error=(
                {
                    "type": "SourceMutationError",
                    "message": "Docker scenario job mutated the host workbook",
                }
                if source_mutated
                else None
            ),
        )

    def _unavailable_trace(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
        before_hash: str,
        started_at: datetime,
        exc: BaseException,
        *,
        identity: DockerRuntimeIdentity | None = None,
    ) -> RuntimeTrace:
        current_hash = _hash_file(workbook)
        runtime_identity = (
            _runtime_identity(identity, {})
            if identity is not None
            else RuntimeIdentity(
                runtime_kind="libreoffice_docker",
                runtime_version="unavailable",
                base_image_digest=BASE_IMAGE_DIGEST,
            )
        )
        return RuntimeTrace(
            trace_id=f"libreoffice-{uuid4().hex}",
            scenario_id=scenario.id,
            runtime_role="target",
            runtime_identity=runtime_identity,
            environment=environment,
            status=GateExecutionStatus.UNAVAILABLE,
            started_at=started_at,
            ended_at=datetime.now(UTC),
            workbook_hash_before=before_hash,
            workbook_hash_after=current_hash,
            error={"type": type(exc).__name__, "message": str(exc)},
        )


def _runtime_identity(identity: DockerRuntimeIdentity, data: dict[str, Any]) -> RuntimeIdentity:
    probe = identity.probe
    return RuntimeIdentity(
        runtime_kind="libreoffice_docker",
        runtime_version=identity.version,
        executable_path=str(probe.get("office_executable") or "") or None,
        executable_sha256=str(probe.get("office_sha256") or "") or None,
        image_reference=identity.image_reference,
        image_digest=identity.image_id,
        base_image_digest=str(probe.get("base_image_digest") or BASE_IMAGE_DIGEST),
        architecture=identity.architecture,
        python_version=str(probe.get("python_version") or "") or None,
        pyuno_identity={
            "uno_module": probe.get("uno_module"),
            "uno_module_sha256": probe.get("uno_module_sha256"),
            "pyuno_native_module": probe.get("pyuno_native_module"),
            "pyuno_native_sha256": probe.get("pyuno_native_sha256"),
            "worker_wrapper": probe.get("worker_wrapper"),
            "worker_wrapper_sha256": probe.get("worker_wrapper_sha256"),
        },
        package_manifest=list(probe.get("installed_package_manifest") or []),
        container_configuration={
            "container_name": data.get("container_name"),
            "job_id": data.get("job_id"),
            "resource_policy": data.get("resource_policy"),
            "exit_code": data.get("container_exit_code"),
            "office_exit_code": (data.get("runtime") or {}).get("office_exit_code")
            if isinstance(data.get("runtime"), dict)
            else None,
        },
        metadata={
            "profile_identifier": (data.get("runtime") or {}).get("profile_identifier")
            if isinstance(data.get("runtime"), dict)
            else None,
            "pipe_name": (data.get("runtime") or {}).get("pipe_name")
            if isinstance(data.get("runtime"), dict)
            else None,
            "final_working_copy_sha256": data.get("final_working_copy_sha256"),
        },
    )


def _runtime_logs(data: dict[str, Any]) -> list[str]:
    logs = [str(item) for item in data.get("logs") or []]
    stderr = str(data.get("container_stderr") or "")
    if stderr:
        logs.append(stderr)
    return logs


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
