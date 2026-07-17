"""Deterministic trace comparison with declared tolerances."""

from __future__ import annotations

import math

from xlsliberator.scenarios.models import (
    ComparisonRules,
    ObservationDifference,
    ObservationValue,
    RuntimeTrace,
    Scenario,
    TraceDiff,
    ValueKind,
)
from xlsliberator.validation_models import GateExecutionStatus


def diff_traces(source: RuntimeTrace, target: RuntimeTrace, scenario: Scenario) -> TraceDiff:
    """Compare matching observation IDs from two executions of one scenario."""
    if source.scenario_id != scenario.id or target.scenario_id != scenario.id:
        raise ValueError("trace scenario IDs do not match the supplied scenario")
    source_steps = {step.step_id: step for step in source.steps}
    target_steps = {step.step_id: step for step in target.steps}
    rules = {
        request.id: request.comparison
        for step in scenario.steps
        for request in (*step.observations_before, *step.observations_after)
    }
    differences: list[ObservationDifference] = []
    for step_id in sorted(set(source_steps) & set(target_steps)):
        source_values = {
            **source_steps[step_id].observations_before,
            **source_steps[step_id].observations_after,
        }
        target_values = {
            **target_steps[step_id].observations_before,
            **target_steps[step_id].observations_after,
        }
        for observation_id in sorted(set(source_values) | set(target_values)):
            source_value = source_values.get(observation_id)
            target_value = target_values.get(observation_id)
            matched, reason = _values_equal(
                source_value, target_value, rules.get(observation_id, ComparisonRules())
            )
            differences.append(
                ObservationDifference(
                    step_id=step_id,
                    observation_id=observation_id,
                    source=source_value,
                    target=target_value,
                    matched=matched,
                    reason=reason,
                )
            )
    missing_source = sorted(set(target_steps) - set(source_steps))
    missing_target = sorted(set(source_steps) - set(target_steps))
    execution_failed = (
        source.status is not GateExecutionStatus.PASSED
        or target.status is not GateExecutionStatus.PASSED
    )
    passed = (
        not execution_failed
        and not missing_source
        and not missing_target
        and all(item.matched for item in differences)
    )
    return TraceDiff(
        source_trace_id=source.trace_id,
        target_trace_id=target.trace_id,
        status=GateExecutionStatus.PASSED if passed else GateExecutionStatus.FAILED,
        differences=differences,
        missing_source_steps=missing_source,
        missing_target_steps=missing_target,
    )


def _values_equal(
    source: ObservationValue | None,
    target: ObservationValue | None,
    rules: ComparisonRules,
) -> tuple[bool, str | None]:
    if source is None or target is None:
        return False, "observation missing from one trace"
    if {source.kind, target.kind} == {ValueKind.EMPTY_CELL, ValueKind.EMPTY_STRING}:
        return (
            (True, None)
            if rules.empty_string_equals_empty_cell
            else (False, "empty string and empty cell are distinct")
        )
    if source.kind is not target.kind:
        return False, f"type differs: {source.kind.value} != {target.kind.value}"
    if source.kind is ValueKind.NUMBER:
        if isinstance(source.value, bool) or isinstance(target.value, bool):
            return False, "boolean values cannot be compared as numbers"
        matched = math.isclose(
            float(source.value),
            float(target.value),
            rel_tol=rules.relative_tolerance,
            abs_tol=rules.absolute_tolerance,
        )
        return matched, None if matched else "numeric values exceed declared tolerance"
    if source.kind is ValueKind.BOOLEAN:
        matched = source.value is target.value
        return matched, None if matched else "boolean values differ"
    if source.kind is ValueKind.ERROR:
        matched = source.error_type == target.error_type
        return matched, None if matched else "formula error identities differ"
    if source.kind in {ValueKind.DATE, ValueKind.DATETIME}:
        matched = (
            source.value == target.value
            and source.date_system == target.date_system
            and source.timezone == target.timezone
        )
        return matched, None if matched else "date value, system, or timezone differs"
    if source.kind is ValueKind.STRING and not rules.string_case_sensitive:
        matched = str(source.value).casefold() == str(target.value).casefold()
    else:
        matched = source.value == target.value
    return matched, None if matched else "normalized values differ"
