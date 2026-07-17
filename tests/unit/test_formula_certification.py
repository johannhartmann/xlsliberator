"""Fail-closed target-parser and differential formula certification tests."""

import hashlib
from datetime import UTC, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

from xlsliberator.formula_certification import certify_formulas
from xlsliberator.formula_semantics import build_formula_ir
from xlsliberator.ir_models import WorkbookIR
from xlsliberator.scenarios.models import (
    Action,
    ActionKind,
    EnvironmentManifest,
    ObservationKind,
    ObservationRequest,
    ObservationValue,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
    ScenarioStep,
    StepResult,
    ValueKind,
)
from xlsliberator.validation_models import (
    GateExecutionStatus,
    SourceRef,
    WorkbookArtifactIR,
)


class FakeParserRuntime:
    def __init__(self, response: dict[str, Any] | None = None) -> None:
        self.response = response or {
            "success": True,
            "data": {
                "tokens": ["1:A1", "2:+", "1:A2"],
                "roundtrip_formula": "=A1+A2",
                "roundtrip_equivalent": True,
                "parser_accepted": True,
                "container_image_id": "sha256:formula-runtime",
            },
        }
        self.calls: list[dict[str, Any]] = []

    def resolve_identity(self, *, probe: bool = True) -> Any:
        assert probe
        return SimpleNamespace(image_id="sha256:formula-runtime")

    def parse_formula(
        self,
        ods_path: Path,
        formula: str,
        *,
        sheet_name: str,
        cell_address: str,
        image_id: str | None = None,
    ) -> dict[str, Any]:
        self.calls.append(
            {
                "ods_path": ods_path,
                "formula": formula,
                "sheet_name": sheet_name,
                "cell_address": cell_address,
                "image_id": image_id,
            }
        )
        return self.response


def _scenario() -> Scenario:
    return Scenario(
        id="formula-runtime",
        steps=[
            ScenarioStep(
                id="recalculate",
                action=Action(kind=ActionKind.RECALCULATE),
                observations_after=[
                    ObservationRequest(
                        id="result",
                        kind=ObservationKind.CELL,
                        selector={"sheet": "Data", "address": "C3"},
                    )
                ],
            )
        ],
    )


def _trace(
    role: str,
    runtime_kind: str,
    observation: ObservationValue,
    *,
    workbook_hash: str = "a" * 64,
    environment: EnvironmentManifest | None = None,
) -> RuntimeTrace:
    now = datetime.now(UTC)
    return RuntimeTrace(
        trace_id=f"{role}-trace",
        scenario_id="formula-runtime",
        runtime_role=role,  # type: ignore[arg-type]
        runtime_identity=RuntimeIdentity(
            runtime_kind=runtime_kind,
            runtime_version="test",
            image_digest="sha256:target" if role == "target" else None,
        ),
        environment=environment or EnvironmentManifest(),
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before=workbook_hash,
        workbook_hash_after=workbook_hash,
        steps=[
            StepResult(
                step_id="recalculate",
                action=ActionKind.RECALCULATE,
                status=GateExecutionStatus.PASSED,
                started_at=now,
                ended_at=now,
                observations_after={"result": observation},
            )
        ],
    )


def _inventory() -> WorkbookArtifactIR:
    source_ref = SourceRef(
        source_file="source.xlsx",
        sheet="Data",
        cell_range="C3",
        artifact_type="formula",
        artifact_id="formula:Data!C3",
    )
    return WorkbookArtifactIR(
        workbook=WorkbookIR(file_path="source.xlsx", file_format="xlsx"),
        formulas=[build_formula_ir(source_ref=source_ref, formula="=A1+A2")],
    )


def _target_file(tmp_path: Path) -> tuple[Path, str]:
    output = tmp_path / "target.ods"
    output.write_bytes(b"target")
    return output, hashlib.sha256(output.read_bytes()).hexdigest()


def test_formula_certification_requires_parser_and_matching_runtime_values(
    tmp_path: Path,
) -> None:
    runtime = FakeParserRuntime()
    output, target_hash = _target_file(tmp_path)
    source = _trace(
        "source",
        "microsoft_excel",
        ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2"),
    )
    target = _trace(
        "target",
        "libreoffice_docker",
        ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2"),
        workbook_hash=target_hash,
    )

    result = certify_formulas(_inventory(), output, source, target, _scenario(), runtime)

    assert result.status is GateExecutionStatus.PASSED
    assert result.records[0].target_parser_result["data"]["roundtrip_equivalent"] is True
    assert result.records[0].source_observation.value == 3
    assert result.records[0].target_observation.value == 3
    assert runtime.calls[0]["image_id"] == "sha256:formula-runtime"


def test_balanced_parentheses_cannot_pass_without_real_traces(tmp_path: Path) -> None:
    runtime = FakeParserRuntime()

    result = certify_formulas(
        _inventory(),
        tmp_path / "target.ods",
        None,
        None,
        _scenario(),
        runtime,
    )

    assert result.status is GateExecutionStatus.FAILED
    assert "Microsoft Excel source trace" in " ".join(result.errors)
    assert "Docker LibreOffice target trace" in " ".join(result.errors)
    assert runtime.calls == []


def test_formula_certification_rejects_parser_roundtrip_change(tmp_path: Path) -> None:
    runtime = FakeParserRuntime(
        {
            "success": True,
            "data": {
                "tokens": ["source"],
                "roundtrip_tokens": ["different"],
                "roundtrip_equivalent": False,
                "parser_accepted": True,
                "container_image_id": "sha256:formula-runtime",
            },
        }
    )
    output, target_hash = _target_file(tmp_path)
    source = _trace(
        "source",
        "microsoft_excel",
        ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2"),
    )
    target = _trace(
        "target",
        "libreoffice_docker",
        ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2"),
        workbook_hash=target_hash,
    )

    result = certify_formulas(_inventory(), output, source, target, _scenario(), runtime)

    assert result.status is GateExecutionStatus.FAILED
    assert "round-trip changed" in result.records[0].errors[0]


def test_formula_certification_preserves_boolean_number_distinction(tmp_path: Path) -> None:
    runtime = FakeParserRuntime()
    output, target_hash = _target_file(tmp_path)
    source = _trace(
        "source",
        "microsoft_excel",
        ObservationValue(kind=ValueKind.BOOLEAN, value=True, formula="=A1+A2"),
    )
    target = _trace(
        "target",
        "libreoffice_docker",
        ObservationValue(kind=ValueKind.NUMBER, value=1, formula="=A1+A2"),
        workbook_hash=target_hash,
    )

    result = certify_formulas(_inventory(), output, source, target, _scenario(), runtime)

    assert result.status is GateExecutionStatus.FAILED
    assert any("type differs" in error for error in result.records[0].errors)


def test_formula_certification_rejects_trace_for_different_target(tmp_path: Path) -> None:
    output, _target_hash = _target_file(tmp_path)
    observation = ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2")
    result = certify_formulas(
        _inventory(),
        output,
        _trace("source", "microsoft_excel", observation),
        _trace("target", "libreoffice_docker", observation, workbook_hash="b" * 64),
        _scenario(),
        FakeParserRuntime(),
    )

    assert result.status is GateExecutionStatus.FAILED
    assert "does not match the target ODS" in " ".join(result.errors)


def test_formula_certification_rejects_environment_drift(tmp_path: Path) -> None:
    output, target_hash = _target_file(tmp_path)
    observation = ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2")
    result = certify_formulas(
        _inventory(),
        output,
        _trace("source", "microsoft_excel", observation),
        _trace(
            "target",
            "libreoffice_docker",
            observation,
            workbook_hash=target_hash,
            environment=EnvironmentManifest(locale="de-DE"),
        ),
        _scenario(),
        FakeParserRuntime(),
    )

    assert result.status is GateExecutionStatus.FAILED
    assert "different environment manifests" in " ".join(result.errors)


def test_formula_certification_rejects_unsupported_formula(tmp_path: Path) -> None:
    output, target_hash = _target_file(tmp_path)
    inventory = _inventory()
    inventory.formulas[0].unsupported_reasons = ["parser coverage unavailable"]
    observation = ObservationValue(kind=ValueKind.NUMBER, value=3, formula="=A1+A2")
    result = certify_formulas(
        inventory,
        output,
        _trace("source", "microsoft_excel", observation),
        _trace("target", "libreoffice_docker", observation, workbook_hash=target_hash),
        _scenario(),
        FakeParserRuntime(),
    )

    assert result.status is GateExecutionStatus.FAILED
    assert "explicitly unsupported" in " ".join(result.records[0].errors)
