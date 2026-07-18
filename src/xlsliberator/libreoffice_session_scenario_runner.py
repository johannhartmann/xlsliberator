"""Acceptance-scenario adapter for the stateful LibreOffice session service."""

from __future__ import annotations

import hashlib
from datetime import UTC, datetime
from pathlib import Path
from typing import Any
from uuid import uuid4

from xlsliberator.docker_runtime import BASE_IMAGE_DIGEST
from xlsliberator.libreoffice_session import DockerSessionBackend, LibreOfficeSessionManager
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    StepResult,
)
from xlsliberator.validation_models import GateExecutionStatus


class LibreOfficeSessionScenarioRunner:
    """Run one exact scenario through a stateful Docker-only service session."""

    def __init__(
        self,
        manager: LibreOfficeSessionManager | None = None,
        *,
        timeout_seconds: int = 120,
    ) -> None:
        self.manager = manager or LibreOfficeSessionManager(
            backend=DockerSessionBackend(timeout_seconds=timeout_seconds)
        )

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> RuntimeTrace:
        source = workbook.resolve()
        source_hash = _hash_file(source)
        started_at = datetime.now(UTC)
        manager = self.manager
        session_id: str | None = None
        session_active = False
        descriptor_data: dict[str, Any] = {}
        result: dict[str, Any] | None = None
        logs: list[str] = []
        attachments: list[str] = []
        try:
            created = manager.create_session(environment=environment)
            session_id = str(created.get("session_id") or "") or None
            descriptor_data = dict(created.get("descriptor") or {})
            if session_id is None:
                result = {
                    "transport_success": False,
                    "operation_status": GateExecutionStatus.FAILED.value,
                    "error": {
                        "type": "MissingSessionId",
                        "message": "session service did not return a session ID",
                    },
                }
            elif not created.get("success"):
                result = created
            else:
                session_active = True
                opened = manager.open_document(session_id, source)
                if not opened.get("success"):
                    result = opened
                else:
                    result = manager.run_scenario(
                        session_id,
                        scenario=scenario,
                        environment=environment,
                    )
        except Exception as exc:
            result = {
                "transport_success": False,
                "operation_status": GateExecutionStatus.UNAVAILABLE.value,
                "error": {"type": type(exc).__name__, "message": str(exc)},
            }
        finally:
            if session_id is not None and session_active:
                try:
                    destroyed = manager.destroy_session(session_id)
                    if not destroyed.get("success"):
                        result = destroyed
                except Exception as exc:
                    result = {
                        "transport_success": False,
                        "operation_status": GateExecutionStatus.FAILED.value,
                        "error": {
                            "type": "SessionCleanupError",
                            "message": f"{type(exc).__name__}: {exc}",
                        },
                    }
            if session_id is not None:
                try:
                    collected = manager.collect_logs(session_id)
                    logs.extend(
                        [
                            str(collected.get("office_log") or ""),
                            str(collected.get("operation_log") or ""),
                        ]
                    )
                    attachments = [
                        str(item) for item in collected.get("attachments") or []
                    ]
                except Exception as exc:
                    logs.append(f"session log collection failed: {type(exc).__name__}: {exc}")

        if result is None:
            result = {
                "transport_success": False,
                "operation_status": GateExecutionStatus.FAILED.value,
                "error": {"type": "MissingSessionResult", "message": "session returned no result"},
            }
        status = _response_status(result)
        try:
            steps = [StepResult.model_validate(item) for item in result.get("steps") or []]
        except (TypeError, ValueError) as exc:
            status = GateExecutionStatus.FAILED
            steps = []
            result["error"] = {"type": "TraceValidationError", "message": str(exc)}
        after_hash = _hash_file(source)
        if after_hash != source_hash:
            status = GateExecutionStatus.FAILED
            result["error"] = {
                "type": "SourceMutationError",
                "message": "stateful session mutated the source workbook",
            }
        runtime = _runtime_identity(descriptor_data, result)
        return RuntimeTrace(
            trace_id=f"libreoffice-session-{uuid4().hex}",
            scenario_id=scenario.id,
            runtime_role="target",
            runtime_identity=runtime,
            environment=environment,
            status=status,
            started_at=_parse_datetime(result.get("started_at"), started_at),
            ended_at=_parse_datetime(result.get("ended_at"), datetime.now(UTC)),
            workbook_hash_before=source_hash,
            workbook_hash_after=after_hash,
            steps=steps,
            attachments=attachments,
            logs=[item for item in logs if item],
            error=(
                dict(result.get("error") or {})
                if status is not GateExecutionStatus.PASSED
                else None
            ),
        )


def _runtime_identity(
    descriptor: dict[str, Any],
    result: dict[str, Any],
) -> RuntimeIdentity:
    return RuntimeIdentity(
        runtime_kind="libreoffice_session_docker",
        runtime_version=str(descriptor.get("libreoffice_version") or "unavailable"),
        executable_path=str(descriptor.get("office_executable") or "") or None,
        executable_sha256=str(descriptor.get("office_executable_sha256") or "") or None,
        image_reference=str(descriptor.get("runtime_image") or "") or None,
        image_digest=str(descriptor.get("runtime_image_digest") or "") or None,
        base_image_digest=BASE_IMAGE_DIGEST,
        container_configuration={
            "session_id": descriptor.get("session_id"),
            "uno_port": descriptor.get("uno_port"),
            "display": descriptor.get("display"),
            "evidence_directory": descriptor.get("evidence_directory"),
        },
        metadata={
            "profile_identifier": descriptor.get("profile_identifier"),
            "working_copy_sha256": result.get("working_copy_sha256"),
        },
    )


def _response_status(response: dict[str, Any]) -> GateExecutionStatus:
    raw = str(response.get("operation_status") or GateExecutionStatus.FAILED.value)
    try:
        return GateExecutionStatus(raw)
    except ValueError:
        return GateExecutionStatus.FAILED


def _parse_datetime(value: Any, fallback: datetime) -> datetime:
    if not value:
        return fallback
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return fallback


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
