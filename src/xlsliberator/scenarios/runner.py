"""Scenario runner interfaces and deterministic fake implementation."""

from __future__ import annotations

import hashlib
from collections.abc import Callable, Mapping
from datetime import UTC, datetime
from typing import Literal, Protocol
from uuid import uuid4

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


class ScenarioRunner(Protocol):
    """Common source/target execution interface."""

    def run(
        self, workbook: bytes, environment: EnvironmentManifest, scenario: Scenario
    ) -> RuntimeTrace:
        """Execute a scenario and return a complete trace."""


class FakeScenarioRunner:
    """Pure runner used to prove serialization and deterministic trace diffing."""

    def __init__(
        self,
        role: Literal["fake_source", "fake_target"],
        observations: Mapping[str, object | ObservationValue],
        *,
        transform: Callable[[object], object] | None = None,
    ) -> None:
        if role not in {"fake_source", "fake_target"}:
            raise ValueError("fake runner role must be fake_source or fake_target")
        self.role = role
        self.observations = dict(observations)
        self.transform = transform or (lambda value: value)

    def run(
        self, workbook: bytes, environment: EnvironmentManifest, scenario: Scenario
    ) -> RuntimeTrace:
        started = datetime.now(UTC)
        step_results = []
        for step in scenario.steps:
            step_started = datetime.now(UTC)
            before = {
                request.id: self._value(request.id, environment)
                for request in step.observations_before
            }
            after = {
                request.id: self._value(request.id, environment)
                for request in step.observations_after
            }
            step_results.append(
                StepResult(
                    step_id=step.id,
                    action=step.action.kind,
                    status=GateExecutionStatus.PASSED,
                    started_at=step_started,
                    ended_at=datetime.now(UTC),
                    observations_before=before,
                    observations_after=after,
                )
            )
        digest = hashlib.sha256(workbook).hexdigest()
        return RuntimeTrace(
            trace_id=f"{self.role}-{uuid4().hex}",
            scenario_id=scenario.id,
            runtime_role=self.role,
            runtime_identity=RuntimeIdentity(runtime_kind=self.role, runtime_version="1"),
            environment=environment,
            status=GateExecutionStatus.PASSED,
            started_at=started,
            ended_at=datetime.now(UTC),
            workbook_hash_before=digest,
            workbook_hash_after=digest,
            steps=step_results,
        )

    def _value(self, observation_id: str, environment: EnvironmentManifest) -> ObservationValue:
        raw = self.observations.get(observation_id)
        if isinstance(raw, ObservationValue):
            return raw
        return normalize_value(
            self.transform(raw),
            date_system=environment.date_system,
            timezone=environment.timezone,
        )
