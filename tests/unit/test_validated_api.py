"""Tests for high-level validated API."""

from datetime import UTC, datetime
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.certification_report import CertificationReport
from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    RuntimeIdentity,
    RuntimeTrace,
    Scenario,
)
from xlsliberator.validated_api import ValidatedTransformationError, transform_validated
from xlsliberator.validation_models import (
    GateExecutionStatus,
    ValidationCertification,
    ValidationGateResult,
)


class _PassingRunner:
    def __init__(self, _plan: object) -> None:
        pass

    def run_all(self) -> CertificationReport:
        return CertificationReport(
            ValidationCertification(
                gate_results=[
                    ValidationGateResult(gate_name="conversion", passed=True, message="ok")
                ]
            )
        )


class _FailingRunner:
    def __init__(self, _plan: object) -> None:
        pass

    def run_all(self) -> CertificationReport:
        return CertificationReport(ValidationCertification(certified=False, errors=["failed gate"]))


def test_transform_validated_returns_report(tmp_path: Path, monkeypatch: Any) -> None:
    """Validated transform should compose convert and ValidationRunner."""
    calls: list[tuple[Path, Path]] = []

    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(
        validated_module,
        "convert",
        lambda input_path, output_path, **_kwargs: calls.append((input_path, output_path)),
    )
    monkeypatch.setattr(validated_module, "ValidationRunner", _PassingRunner)

    report = transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods")

    assert report.certification.certified
    assert calls == [(tmp_path / "in.xlsx", tmp_path / "out.ods")]


def test_transform_validated_strict_failure_raises(tmp_path: Path, monkeypatch: Any) -> None:
    """Strict validated transform should raise with evidence report."""
    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(validated_module, "convert", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(validated_module, "ValidationRunner", _FailingRunner)

    with pytest.raises(ValidatedTransformationError) as exc:
        transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods", strict=True)

    assert not exc.value.report.certification.certified


def test_transform_validated_non_strict_returns_failed_report(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Non-strict validated transform should return failed certification evidence."""
    import xlsliberator.validated_api as validated_module

    monkeypatch.setattr(validated_module, "convert", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(validated_module, "ValidationRunner", _FailingRunner)

    report = transform_validated(tmp_path / "in.xlsx", tmp_path / "out.ods", strict=False)

    assert not report.certification.certified


def test_transform_validated_runs_docker_target_scenario_from_source_evidence(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """A supplied source trace should drive the exact target scenario after conversion."""
    import xlsliberator.validated_api as validated_module

    now = datetime.now(UTC)
    scenario = Scenario(id="formula", steps=[])
    environment = EnvironmentManifest(locale="de-DE")
    source_trace = RuntimeTrace(
        trace_id="source",
        scenario_id=scenario.id,
        runtime_role="source",
        runtime_identity=RuntimeIdentity(
            runtime_kind="microsoft_excel",
            runtime_version="test",
        ),
        environment=environment,
        status=GateExecutionStatus.PASSED,
        started_at=now,
        ended_at=now,
        workbook_hash_before="a" * 64,
        workbook_hash_after="a" * 64,
    )
    target_trace = source_trace.model_copy(
        update={
            "trace_id": "target",
            "runtime_role": "target",
            "runtime_identity": RuntimeIdentity(
                runtime_kind="libreoffice_docker",
                runtime_version="26.2.4.2",
                image_digest="sha256:fixed",
            ),
        }
    )
    captured: dict[str, Any] = {}

    class _ScenarioRunner:
        def run(
            self,
            output: Path,
            supplied_environment: EnvironmentManifest,
            supplied_scenario: Scenario,
        ) -> RuntimeTrace:
            captured["scenario_call"] = (output, supplied_environment, supplied_scenario)
            return target_trace

    class _CapturingRunner(_PassingRunner):
        def __init__(self, plan: object) -> None:
            captured["plan"] = plan

    monkeypatch.setattr(validated_module, "convert", lambda *_args, **_kwargs: None)
    monkeypatch.setattr(validated_module, "LibreOfficeScenarioRunner", _ScenarioRunner)
    monkeypatch.setattr(validated_module, "ValidationRunner", _CapturingRunner)
    output = tmp_path / "out.ods"

    transform_validated(
        tmp_path / "in.xlsx",
        output,
        scenario=scenario,
        source_trace=source_trace,
    )

    assert captured["scenario_call"] == (output, environment, scenario)
    plan = captured["plan"]
    assert plan.target_trace == target_trace
    assert "target_scenario" in plan.enabled_gates
