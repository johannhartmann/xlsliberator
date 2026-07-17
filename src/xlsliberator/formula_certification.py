"""Fail-closed formula certification from parser and differential runtime evidence."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Literal, Protocol

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.scenarios.diff import diff_traces
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    ObservationValue,
    RuntimeTrace,
    Scenario,
)
from xlsliberator.validation_models import (
    FormulaIR,
    GateExecutionStatus,
    SourceRef,
    TargetRef,
    WorkbookArtifactIR,
)


class FormulaParserRuntime(Protocol):
    """Docker parser boundary required by formula certification."""

    def resolve_identity(self, *, probe: bool = True) -> Any:
        """Resolve the immutable Docker runtime identity."""

    def parse_formula(
        self,
        ods_path: Path,
        formula: str,
        *,
        sheet_name: str,
        cell_address: str,
        image_id: str | None = None,
    ) -> dict[str, Any]:
        """Parse a formula inside the target document context."""


class FormulaEvidenceRecord(BaseModel):
    """Complete evidence for one source formula and its target outcome."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    source_ref: SourceRef
    target_ref: TargetRef | None = None
    source_formula: str
    target_formula: str | None = None
    target_parser_result: dict[str, Any] = Field(default_factory=dict)
    source_observation: ObservationValue | None = None
    target_observation: ObservationValue | None = None
    observation_id: str | None = None
    trace_difference: dict[str, Any] | None = None
    dependencies: list[str] = Field(default_factory=list)
    spill_context: dict[str, Any] = Field(default_factory=dict)
    calculation_settings: dict[str, Any] = Field(default_factory=dict)
    calculation_order: dict[str, Any] = Field(default_factory=dict)
    semantic_diagnostics: list[str] = Field(default_factory=list)
    runtime_evidence_requirements: list[str] = Field(default_factory=list)
    unsupported_reasons: list[str] = Field(default_factory=list)
    status: GateExecutionStatus
    errors: list[str] = Field(default_factory=list)


class FormulaCertificationResult(BaseModel):
    """Aggregate result whose passed state requires every formula record to pass."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    status: GateExecutionStatus
    source_trace_id: str | None = None
    target_trace_id: str | None = None
    runtime_image_id: str | None = None
    formula_count: int
    records: list[FormulaEvidenceRecord] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)


def certify_formulas(
    inventory: WorkbookArtifactIR,
    target_path: Path,
    source_trace: RuntimeTrace | None,
    target_trace: RuntimeTrace | None,
    scenario: Scenario | None,
    runtime: FormulaParserRuntime,
) -> FormulaCertificationResult:
    """Require target parsing plus exact source/target execution evidence per formula."""
    formulas = list(inventory.formulas)
    if not formulas:
        return FormulaCertificationResult(
            status=GateExecutionStatus.PASSED,
            formula_count=0,
        )
    preflight_errors = _trace_preflight_errors(
        inventory,
        target_path,
        source_trace,
        target_trace,
        scenario,
    )
    if preflight_errors:
        return FormulaCertificationResult(
            status=GateExecutionStatus.FAILED,
            source_trace_id=source_trace.trace_id if source_trace else None,
            target_trace_id=target_trace.trace_id if target_trace else None,
            formula_count=len(formulas),
            errors=preflight_errors,
        )
    assert source_trace is not None
    assert target_trace is not None
    assert scenario is not None
    trace_diff = diff_traces(source_trace, target_trace, scenario)
    observation_requests = _cell_observation_requests(scenario)
    source_observations = _trace_observations(source_trace)
    target_observations = _trace_observations(target_trace)
    differences = {(item.step_id, item.observation_id): item for item in trace_diff.differences}
    try:
        identity = runtime.resolve_identity(probe=True)
        image_id = str(identity.image_id)
    except Exception as exc:
        return FormulaCertificationResult(
            status=GateExecutionStatus.FAILED,
            source_trace_id=source_trace.trace_id,
            target_trace_id=target_trace.trace_id,
            formula_count=len(formulas),
            errors=[f"Docker FormulaParser runtime unavailable: {exc}"],
        )

    records = [
        _certify_one_formula(
            formula,
            target_path,
            observation_requests,
            source_observations,
            target_observations,
            differences,
            source_trace.environment,
            runtime,
            image_id,
        )
        for formula in formulas
    ]
    failed = [record for record in records if record.status is not GateExecutionStatus.PASSED]
    return FormulaCertificationResult(
        status=GateExecutionStatus.FAILED if failed else GateExecutionStatus.PASSED,
        source_trace_id=source_trace.trace_id,
        target_trace_id=target_trace.trace_id,
        runtime_image_id=image_id,
        formula_count=len(formulas),
        records=records,
        errors=[error for record in failed for error in record.errors],
    )


def _certify_one_formula(
    formula: FormulaIR,
    target_path: Path,
    observation_requests: dict[tuple[str, str], tuple[str, str]],
    source_observations: dict[tuple[str, str], ObservationValue],
    target_observations: dict[tuple[str, str], ObservationValue],
    differences: dict[tuple[str, str], Any],
    environment: EnvironmentManifest,
    runtime: FormulaParserRuntime,
    image_id: str,
) -> FormulaEvidenceRecord:
    source_ref = formula.source_ref
    errors = _formula_environment_errors(formula, environment)
    if formula.unsupported_reasons:
        errors.extend(
            f"formula is explicitly unsupported: {reason}" for reason in formula.unsupported_reasons
        )
    sheet = source_ref.sheet
    address = source_ref.cell_range
    if not sheet or not address:
        errors.append("formula has no sheet/cell context and cannot receive execution evidence")
        return _record(formula, errors=errors)
    request_key = observation_requests.get((sheet, address))
    if request_key is None:
        errors.append(f"scenario does not observe formula cell {sheet}!{address}")
        return _record(formula, errors=errors)
    source_observation = source_observations.get(request_key)
    target_observation = target_observations.get(request_key)
    difference = differences.get(request_key)
    target_formula = target_observation.formula if target_observation else None
    if source_observation is None:
        errors.append("source trace omitted the formula observation")
    elif not source_observation.formula:
        errors.append("source trace observation omitted the executed source formula")
    if target_observation is None:
        errors.append("target trace omitted the formula observation")
    elif not target_formula:
        errors.append("target trace observation omitted the executed target formula")
    parser_result: dict[str, Any] = {}
    if target_formula:
        try:
            parser_result = runtime.parse_formula(
                target_path,
                target_formula,
                sheet_name=sheet,
                cell_address=address,
                image_id=image_id,
            )
        except Exception as exc:
            errors.append(f"Docker FormulaParser request failed: {exc}")
        else:
            if not parser_result.get("success"):
                message = str((parser_result.get("error") or {}).get("message") or "parse failed")
                errors.append(f"target FormulaParser rejected formula: {message}")
            data = dict(parser_result.get("data") or {})
            parser_image_id = str(data.get("container_image_id") or "")
            if parser_result.get("success") and parser_image_id != image_id:
                errors.append("target FormulaParser response has the wrong runtime image identity")
            if parser_result.get("success") and data.get("parser_accepted") is not True:
                errors.append("target FormulaParser recovered from invalid formula syntax")
            if parser_result.get("success") and not data.get("roundtrip_equivalent"):
                errors.append("target FormulaParser round-trip changed the token stream")
            if parser_result.get("success") and not data.get("tokens"):
                errors.append("target FormulaParser returned no tokens")
    if difference is None:
        errors.append("source/target trace diff omitted the formula observation")
    elif not bool(difference.matched):
        errors.append(f"runtime value/error differs: {difference.reason or 'unspecified'}")
    target_ref = TargetRef(
        target_file=str(target_path),
        sheet=sheet,
        cell_range=address,
        artifact_type="formula",
        artifact_id=f"formula:{sheet}!{address}",
    )
    formula.target_ref = target_ref
    formula.target_formula_text = target_formula
    return FormulaEvidenceRecord(
        source_ref=source_ref,
        target_ref=target_ref,
        source_formula=formula.original_formula_text or formula.formula_text,
        target_formula=target_formula,
        target_parser_result=parser_result,
        source_observation=source_observation,
        target_observation=target_observation,
        observation_id=request_key[1],
        trace_difference=(difference.model_dump(mode="json") if difference else None),
        dependencies=formula.dependencies,
        spill_context=formula.array_metadata,
        calculation_settings=formula.calculation_settings,
        calculation_order=formula.calculation_order,
        semantic_diagnostics=formula.semantic_diagnostics,
        runtime_evidence_requirements=formula.runtime_evidence_requirements,
        unsupported_reasons=formula.unsupported_reasons,
        status=GateExecutionStatus.FAILED if errors else GateExecutionStatus.PASSED,
        errors=errors,
    )


def _record(formula: FormulaIR, *, errors: list[str]) -> FormulaEvidenceRecord:
    return FormulaEvidenceRecord(
        source_ref=formula.source_ref,
        source_formula=formula.original_formula_text or formula.formula_text,
        dependencies=formula.dependencies,
        spill_context=formula.array_metadata,
        calculation_settings=formula.calculation_settings,
        calculation_order=formula.calculation_order,
        semantic_diagnostics=formula.semantic_diagnostics,
        runtime_evidence_requirements=formula.runtime_evidence_requirements,
        unsupported_reasons=formula.unsupported_reasons,
        status=GateExecutionStatus.FAILED,
        errors=errors,
    )


def _trace_preflight_errors(
    inventory: WorkbookArtifactIR,
    target_path: Path,
    source: RuntimeTrace | None,
    target: RuntimeTrace | None,
    scenario: Scenario | None,
) -> list[str]:
    errors: list[str] = []
    if scenario is None:
        errors.append("formula certification requires the exact scenario definition")
    if source is None:
        errors.append("formula certification requires a Microsoft Excel source trace")
    elif source.runtime_identity.runtime_kind != "microsoft_excel":
        errors.append("source trace is not evidence from Microsoft Excel")
    elif source.runtime_role != "source":
        errors.append("source trace runtime role is not source")
    elif source.status is not GateExecutionStatus.PASSED:
        errors.append(f"source trace status is {source.status.value}")
    if target is None:
        errors.append("formula certification requires a Docker LibreOffice target trace")
    elif target.runtime_identity.runtime_kind != "libreoffice_docker":
        errors.append("target trace is not evidence from Docker LibreOffice")
    elif target.runtime_role != "target":
        errors.append("target trace runtime role is not target")
    elif not target.runtime_identity.image_digest:
        errors.append("target trace omits the immutable Docker image digest")
    elif target.status is not GateExecutionStatus.PASSED:
        errors.append(f"target trace status is {target.status.value}")
    if scenario and source and source.scenario_id != scenario.id:
        errors.append("source trace scenario does not match the supplied scenario")
    if scenario and target and target.scenario_id != scenario.id:
        errors.append("target trace scenario does not match the supplied scenario")
    if source and target and source.environment != target.environment:
        errors.append("source and target traces use different environment manifests")
    if (
        source
        and inventory.source_sha256
        and source.workbook_hash_before != inventory.source_sha256
    ):
        errors.append("source trace workbook hash does not match the source inventory")
    if target:
        if not target_path.is_file():
            errors.append("target workbook is missing")
        else:
            target_hash = _hash_file(target_path)
            if target.workbook_hash_before != target_hash:
                errors.append("target trace workbook hash does not match the target ODS")
            if target.workbook_hash_after != target_hash:
                errors.append("target trace does not prove the target ODS remained immutable")
    if scenario:
        for step in scenario.steps:
            for request in (*step.observations_before, *step.observations_after):
                if (
                    request.kind.value == "cell"
                    and request.comparison.empty_string_equals_empty_cell
                ):
                    errors.append(
                        f"formula observation {request.id} enables the forbidden empty-cell shortcut"
                    )
    return errors


def _formula_environment_errors(
    formula: FormulaIR,
    environment: EnvironmentManifest,
) -> list[str]:
    errors: list[str] = []
    if "external_reference" in formula.semantic_features and not environment.external_workbooks:
        errors.append("external formula reference is not declared in the environment manifest")
    settings = formula.calculation_settings
    mode = settings.get("mode")
    normalized_modes = {
        "auto": "automatic",
        "manual": "manual",
        "autoNoTable": "automatic_except_tables",
    }
    if mode and normalized_modes.get(str(mode), str(mode)) != environment.calculation_mode:
        errors.append("formula calculation mode differs from the execution environment")
    iterate = settings.get("iterate")
    if iterate is not None and bool(iterate) != environment.iterative_calculation:
        errors.append("formula iterative-calculation setting differs from the environment")
    return errors


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cell_observation_requests(scenario: Scenario) -> dict[tuple[str, str], tuple[str, str]]:
    observation_requests: dict[tuple[str, str], tuple[str, str]] = {}
    for step in scenario.steps:
        for request in (*step.observations_before, *step.observations_after):
            if request.kind.value != "cell":
                continue
            sheet = str(request.selector.get("sheet") or request.selector.get("sheet_name") or "")
            address = str(
                request.selector.get("address") or request.selector.get("cell_address") or ""
            )
            if sheet and address:
                observation_requests[(sheet, address)] = (step.id, request.id)
    return observation_requests


def _trace_observations(trace: RuntimeTrace) -> dict[tuple[str, str], ObservationValue]:
    result: dict[tuple[str, str], ObservationValue] = {}
    for step in trace.steps:
        for observation_id, value in {
            **step.observations_before,
            **step.observations_after,
        }.items():
            result[(step.step_id, observation_id)] = value
    return result
