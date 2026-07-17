"""Deterministic VBA micro-conformance corpus and Windows source-trace runner."""

from __future__ import annotations

from pathlib import Path

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.excel_oracle import ExcelOracle
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    Scenario,
    ScenarioStep,
)
from xlsliberator.validation_models import GateExecutionStatus


class VBAMicroProgram(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str = "1.0.0"
    program_id: str
    feature: str
    source_code: str
    procedure_name: str
    workbook_fixture: str
    scenario: Scenario
    required_capabilities: list[str] = Field(default_factory=list)


class VBAMicroTraceResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    program_id: str
    status: GateExecutionStatus
    trace_path: str | None = None
    real_excel_trace: bool = False
    error: dict[str, str] | None = None


def default_vba_micro_programs() -> list[VBAMicroProgram]:
    """Return the source corpus covering the required semantic categories."""
    sources = {
        "type-coercion": 'Range("A1").Value = Empty + True',
        "byref": "Increment value",
        "default-properties": 'Range("A1") = Range("B1")',
        "arrays": "ReDim Preserve values(0 To 3)",
        "error-handling": "On Error Resume Next",
        "classes": "Set value = New Counter",
        "events": "RaiseEvent Changed(Cancel)",
        "range-operations": 'Range("A1:B2").Value = Array(1, 2, 3, 4)',
    }
    programs = []
    for feature, statement in sources.items():
        procedure = "Conformance_" + feature.replace("-", "_")
        programs.append(
            VBAMicroProgram(
                program_id=f"vba-{feature}-v1",
                feature=feature,
                source_code=f"Public Sub {procedure}()\n    {statement}\nEnd Sub",
                procedure_name=procedure,
                workbook_fixture=f"{feature}.xlsm",
                scenario=_micro_scenario(feature, procedure),
                required_capabilities=["macro_execution"],
            )
        )
    return programs


def generate_windows_micro_traces(
    oracle: ExcelOracle,
    fixture_dir: Path,
    output_dir: Path,
    environment: EnvironmentManifest,
    programs: list[VBAMicroProgram] | None = None,
) -> list[VBAMicroTraceResult]:
    """Generate real Excel traces when both oracle and fixture are available."""
    output_dir.mkdir(parents=True, exist_ok=True)
    results: list[VBAMicroTraceResult] = []
    for program in programs or default_vba_micro_programs():
        workbook = fixture_dir / program.workbook_fixture
        if not workbook.is_file():
            results.append(
                VBAMicroTraceResult(
                    program_id=program.program_id,
                    status=GateExecutionStatus.UNAVAILABLE,
                    error={"type": "missing_fixture", "message": str(workbook)},
                )
            )
            continue
        outcome = oracle.run(workbook, environment, program.scenario)
        trace = outcome.trace
        is_real = bool(
            trace
            and trace.runtime_role == "source"
            and trace.runtime_identity.runtime_kind == "microsoft_excel"
        )
        trace_path = None
        if is_real and trace is not None:
            destination = output_dir / f"{program.program_id}.source-trace.json"
            destination.write_text(trace.model_dump_json(indent=2), encoding="utf-8")
            trace_path = str(destination)
        error = None
        if outcome.error:
            error = {str(key): str(value) for key, value in outcome.error.items()}
        elif trace is not None and not is_real:
            error = {
                "type": "non_excel_trace",
                "message": trace.runtime_identity.runtime_kind,
            }
        results.append(
            VBAMicroTraceResult(
                program_id=program.program_id,
                status=outcome.status if is_real else GateExecutionStatus.UNAVAILABLE,
                trace_path=trace_path,
                real_excel_trace=is_real,
                error=error,
            )
        )
    return results


def _micro_scenario(feature: str, procedure: str) -> Scenario:
    return Scenario(
        id=f"vba-micro-{feature}-v1",
        description=f"Microsoft Excel source trace for VBA {feature}",
        steps=[
            ScenarioStep(id="open", action=Action(kind=ActionKind.OPEN)),
            ScenarioStep(
                id="invoke",
                action=Action(
                    kind=ActionKind.INVOKE_MACRO,
                    parameters={"macro_name": procedure},
                ),
                observations_after=[
                    ObservationRequest(
                        id="result",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Sheet1", "address": "A1"},
                    )
                ],
            ),
            ScenarioStep(id="save", action=Action(kind=ActionKind.SAVE)),
            ScenarioStep(id="reopen", action=Action(kind=ActionKind.REOPEN)),
        ],
    )
