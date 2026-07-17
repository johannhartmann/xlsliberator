"""Microsoft Excel source-oracle protocol and secure remote client."""

from __future__ import annotations

import base64
import json
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Mapping
from datetime import UTC, datetime
from pathlib import Path
from typing import Any, Protocol
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, model_validator

from xlsliberator.execution_sandbox import (
    ExecutionKind,
    SandboxBackendKind,
    SandboxPolicy,
)
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    ObservationValue,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    StepResult,
)
from xlsliberator.scenarios.normalize import normalize_value
from xlsliberator.validation_models import GateExecutionStatus

ORACLE_PROTOCOL_VERSION = "1.0.0"


class OracleProtocolError(RuntimeError):
    """A source-oracle response violated the versioned protocol."""


class OracleTransportUnavailable(RuntimeError):
    """The separately secured Windows oracle cannot be reached."""


class OracleRunResult(BaseModel):
    """Source execution result with explicit unavailable/failure states."""

    model_config = ConfigDict(extra="forbid")

    schema_version: str = ORACLE_PROTOCOL_VERSION
    status: GateExecutionStatus
    trace: RuntimeTrace | None = None
    attachments: dict[str, str] = Field(default_factory=dict)
    error: dict[str, Any] | None = None

    @model_validator(mode="after")
    def passed_requires_trace(self) -> OracleRunResult:
        if self.status is GateExecutionStatus.PASSED and self.trace is None:
            raise ValueError("a passed Excel oracle response requires a runtime trace")
        return self

    @property
    def succeeded(self) -> bool:
        return self.status is GateExecutionStatus.PASSED and self.trace is not None


class ExcelOracle(Protocol):
    """Source runtime contract consumed by differential orchestration."""

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> OracleRunResult:
        """Execute a workbook scenario in licensed Microsoft Excel."""


class OracleTransport(Protocol):
    """Transport to a trusted Windows worker or worker service."""

    def submit(self, request: Mapping[str, Any], timeout_seconds: float) -> Mapping[str, Any]:
        """Submit one request and return one decoded response."""


class HTTPJSONLinesOracleTransport:
    """HTTPS JSON-lines transport for a separately secured Windows service."""

    def __init__(self, endpoint: str, *, bearer_token: str | None = None) -> None:
        parsed = urllib.parse.urlparse(endpoint)
        if parsed.scheme != "https" and parsed.hostname not in {"localhost", "127.0.0.1", "::1"}:
            raise ValueError("remote Excel oracle endpoints must use HTTPS")
        self.endpoint = endpoint
        self.bearer_token = bearer_token

    def submit(self, request: Mapping[str, Any], timeout_seconds: float) -> Mapping[str, Any]:
        body = (json.dumps(request, separators=(",", ":")) + "\n").encode("utf-8")
        headers = {"Content-Type": "application/x-ndjson"}
        if self.bearer_token:
            headers["Authorization"] = f"Bearer {self.bearer_token}"
        http_request = urllib.request.Request(
            self.endpoint, data=body, headers=headers, method="POST"
        )
        try:
            # The constructor rejects non-HTTPS remote endpoints; HTTP is permitted only for
            # loopback workers used by the isolated oracle harness.
            with urllib.request.urlopen(  # nosec B310
                http_request, timeout=timeout_seconds
            ) as response:
                raw = response.read().decode("utf-8")
        except (OSError, TimeoutError, urllib.error.URLError) as exc:
            raise OracleTransportUnavailable(f"Windows Excel oracle is unavailable: {exc}") from exc
        lines = [line for line in raw.splitlines() if line.strip()]
        if len(lines) != 1:
            raise OracleProtocolError("Excel oracle must return exactly one JSON response line")
        try:
            payload = json.loads(lines[0])
        except json.JSONDecodeError as exc:
            raise OracleProtocolError("Excel oracle returned malformed JSON") from exc
        if not isinstance(payload, dict):
            raise OracleProtocolError("Excel oracle response must be an object")
        return payload


class WindowsExcelOracleClient:
    """Linux/Docker-side client; Microsoft Excel itself remains on Windows."""

    def __init__(
        self,
        transport: OracleTransport,
        *,
        timeout_seconds: float = 180.0,
        sandbox_policy: SandboxPolicy | None = None,
    ) -> None:
        self.transport = transport
        self.timeout_seconds = timeout_seconds
        self.sandbox_policy = sandbox_policy or SandboxPolicy(
            backend=SandboxBackendKind.REMOTE_WORKER
        )

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> OracleRunResult:
        if not workbook.is_file():
            return OracleRunResult(
                status=GateExecutionStatus.FAILED,
                error={"type": "missing_workbook", "message": str(workbook)},
            )
        request = {
            "schema_version": ORACLE_PROTOCOL_VERSION,
            "request_id": uuid4().hex,
            "workbook_name": workbook.name,
            "workbook_base64": base64.b64encode(workbook.read_bytes()).decode("ascii"),
            "environment": environment.model_dump(mode="json"),
            "scenario": scenario.model_dump(mode="json"),
            "timeout_seconds": self.timeout_seconds,
            "execution_kind": ExecutionKind.SOURCE_ORACLE.value,
            "sandbox_policy": self.sandbox_policy.model_dump(mode="json"),
        }
        try:
            raw = self.transport.submit(request, self.timeout_seconds)
            result = OracleRunResult.model_validate(raw)
        except OracleTransportUnavailable as exc:
            return OracleRunResult(
                status=GateExecutionStatus.UNAVAILABLE,
                error={"type": "oracle_unavailable", "message": str(exc)},
            )
        except Exception as exc:
            return OracleRunResult(
                status=GateExecutionStatus.FAILED,
                error={"type": "oracle_protocol_error", "message": str(exc)},
            )
        if result.trace and (
            result.trace.runtime_role != "source"
            or result.trace.runtime_identity.runtime_kind != "microsoft_excel"
        ):
            return OracleRunResult(
                status=GateExecutionStatus.FAILED,
                error={
                    "type": "invalid_runtime_identity",
                    "message": (
                        f"{result.trace.runtime_role}:{result.trace.runtime_identity.runtime_kind}"
                    ),
                },
            )
        if result.trace:
            raw_policy = result.trace.runtime_identity.container_configuration.get("sandbox_policy")
            try:
                attested = SandboxPolicy.model_validate(raw_policy)
            except Exception:
                return OracleRunResult(
                    status=GateExecutionStatus.FAILED,
                    error={
                        "type": "missing_sandbox_attestation",
                        "message": "Excel source trace lacks a valid remote sandbox attestation",
                    },
                )
            if attested.backend not in {
                SandboxBackendKind.REMOTE_WORKER,
                SandboxBackendKind.MICROVM,
            }:
                return OracleRunResult(
                    status=GateExecutionStatus.FAILED,
                    error={
                        "type": "invalid_sandbox_backend",
                        "message": "Excel source execution requires a remote worker or microVM",
                    },
                )
        return result


class UnavailableExcelOracle:
    """Explicit source-oracle state for hosts without a configured Windows worker."""

    def __init__(self, reason: str = "Windows Excel oracle is not configured") -> None:
        self.reason = reason

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> OracleRunResult:
        del workbook, environment, scenario
        return OracleRunResult(
            status=GateExecutionStatus.UNAVAILABLE,
            error={"type": "oracle_unavailable", "message": self.reason},
        )


class FakeExcelOracle:
    """Deterministic fake that exercises every action without claiming Excel ran."""

    def __init__(
        self,
        observations: Mapping[str, object | ObservationValue] | None = None,
        *,
        action_statuses: Mapping[str, GateExecutionStatus] | None = None,
    ) -> None:
        self.observations = dict(observations or {})
        self.action_statuses = dict(action_statuses or {})

    def run(
        self,
        workbook: Path,
        environment: EnvironmentManifest,
        scenario: Scenario,
    ) -> OracleRunResult:
        started = datetime.now(UTC)
        steps: list[StepResult] = []
        for step in scenario.steps:
            status = self.action_statuses.get(step.action.kind.value, GateExecutionStatus.PASSED)
            before = {
                request.id: self._observation(request.id, environment)
                for request in step.observations_before
            }
            after = {
                request.id: self._observation(request.id, environment)
                for request in step.observations_after
            }
            steps.append(
                StepResult(
                    step_id=step.id,
                    action=step.action.kind,
                    status=status,
                    started_at=started,
                    ended_at=datetime.now(UTC),
                    observations_before=before,
                    observations_after=after,
                    error=(
                        None
                        if status is GateExecutionStatus.PASSED
                        else {"type": status.value, "message": "configured fake outcome"}
                    ),
                )
            )
        required_failed = any(
            result.status is not GateExecutionStatus.PASSED
            and next(step for step in scenario.steps if step.id == result.step_id).action.required
            for result in steps
        )
        digest = _sha256_file(workbook)
        trace = RuntimeTrace(
            trace_id=f"fake-excel-{uuid4().hex}",
            scenario_id=scenario.id,
            runtime_role="source",
            runtime_identity=RuntimeIdentity(
                runtime_kind="fake_excel_oracle", runtime_version="1.0.0"
            ),
            environment=environment,
            status=(GateExecutionStatus.FAILED if required_failed else GateExecutionStatus.PASSED),
            started_at=started,
            ended_at=datetime.now(UTC),
            workbook_hash_before=digest,
            workbook_hash_after=digest,
            steps=steps,
        )
        return OracleRunResult(status=trace.status, trace=trace)

    def _observation(
        self, observation_id: str, environment: EnvironmentManifest
    ) -> ObservationValue:
        value = self.observations.get(observation_id)
        if isinstance(value, ObservationValue):
            return value
        return normalize_value(
            value, date_system=environment.date_system, timezone=environment.timezone
        )


def load_source_trace_fixture(path: Path) -> RuntimeTrace:
    """Load a recorded Excel trace without representing it as local execution."""
    trace = RuntimeTrace.model_validate_json(path.read_text(encoding="utf-8"))
    if trace.runtime_role != "source" or trace.runtime_identity.runtime_kind != "microsoft_excel":
        raise ValueError("source trace fixture must identify a real Microsoft Excel source runtime")
    if not trace.runtime_identity.runtime_version:
        raise ValueError("source trace fixture omits the Excel build identity")
    return trace


def _sha256_file(path: Path) -> str:
    import hashlib

    return hashlib.sha256(path.read_bytes()).hexdigest()
