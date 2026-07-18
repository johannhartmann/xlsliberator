"""Fail-closed public acceptance assertion evaluation."""

from __future__ import annotations

from xlsliberator.scenarios.diff import compare_values
from xlsliberator.scenarios.models import (
    AcceptanceDefinition,
    AcceptanceEvaluation,
    AssertionResult,
    ObservationRequest,
    ObservationValue,
    RuntimeTrace,
)
from xlsliberator.validation_models import GateExecutionStatus


def evaluate_trace(
    acceptance: AcceptanceDefinition,
    trace: RuntimeTrace,
) -> AcceptanceEvaluation:
    """Evaluate required actions and authored expectations against one trace."""
    scenario = acceptance.scenario
    if trace.scenario_id != scenario.id:
        raise ValueError("trace scenario ID does not match the acceptance definition")

    trace_steps = {step.step_id: step for step in trace.steps}
    action_statuses: dict[str, GateExecutionStatus] = {}
    assertions: list[AssertionResult] = []
    required_failures: list[str] = []

    for step in scenario.steps:
        result = trace_steps.get(step.id)
        status = result.status if result is not None else GateExecutionStatus.NOT_RUN
        action_statuses[step.id] = status
        if step.action.required and status is not GateExecutionStatus.PASSED:
            required_failures.append(f"action {step.id} was {status.value}")

        before = result.observations_before if result is not None else {}
        after = result.observations_after if result is not None else {}
        assertions.extend(
            _evaluate_observations(step.id, "before", step.observations_before, before)
        )
        assertions.extend(_evaluate_observations(step.id, "after", step.observations_after, after))

    required_failures.extend(
        f"observation {item.observation_id} {item.reason or 'failed'}"
        for item in assertions
        if item.required and item.status is not GateExecutionStatus.PASSED
    )
    if trace.status is not GateExecutionStatus.PASSED:
        required_failures.append(f"runtime trace was {trace.status.value}")

    return AcceptanceEvaluation(
        migration_id=acceptance.migration.id,
        scenario_id=scenario.id,
        trace_id=trace.trace_id,
        status=(
            GateExecutionStatus.PASSED if not required_failures else GateExecutionStatus.FAILED
        ),
        action_statuses=action_statuses,
        assertions=assertions,
        required_failures=required_failures,
    )


def _evaluate_observations(
    step_id: str,
    phase: str,
    requests: list[ObservationRequest],
    actual: dict[str, ObservationValue],
) -> list[AssertionResult]:
    results: list[AssertionResult] = []
    for request in requests:
        value = actual.get(request.id)
        if value is None:
            status = GateExecutionStatus.FAILED if request.required else GateExecutionStatus.SKIPPED
            results.append(
                AssertionResult(
                    step_id=step_id,
                    observation_id=request.id,
                    phase="before" if phase == "before" else "after",
                    required=request.required,
                    status=status,
                    expected=request.expected,
                    reason="was not observed",
                )
            )
            continue
        if request.expected is None:
            results.append(
                AssertionResult(
                    step_id=step_id,
                    observation_id=request.id,
                    phase="before" if phase == "before" else "after",
                    required=request.required,
                    status=GateExecutionStatus.PASSED,
                    actual=value,
                    reason="evidence-only observation",
                )
            )
            continue
        matched, reason = compare_values(request.expected, value, request.comparison)
        results.append(
            AssertionResult(
                step_id=step_id,
                observation_id=request.id,
                phase="before" if phase == "before" else "after",
                required=request.required,
                status=(GateExecutionStatus.PASSED if matched else GateExecutionStatus.FAILED),
                expected=request.expected,
                actual=value,
                reason=reason,
            )
        )
    return results
