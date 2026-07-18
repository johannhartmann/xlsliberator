"""Cross-repository agent benchmark release-gate tests."""

from __future__ import annotations

from pathlib import Path

import pytest
from pydantic import ValidationError

from xlsliberator.agent_evaluation import (
    AgentBenchmarkReport,
    AgentEvaluationStatus,
    AgentEvaluatorName,
    AgentEvaluatorResult,
    AgentMigrationEvaluation,
    AgentPartitionSummary,
    require_agent_benchmark_release,
)


def _summary(
    partition: str,
    *,
    passed: int,
    failed: int = 0,
    skipped: int = 0,
    unavailable: int = 0,
    not_run: int = 0,
) -> AgentPartitionSummary:
    decisive = passed + failed
    return AgentPartitionSummary.model_validate(
        {
            "partition": partition,
            "counts": {
                "passed": passed,
                "failed": failed,
                "skipped": skipped,
                "unavailable": unavailable,
                "not_run": not_run,
            },
            "decisive_pass_rate": passed / decisive if decisive else None,
            "hidden_definitions_included": False,
        }
    )


def _evaluators(
    *,
    failed: AgentEvaluatorName | None = None,
) -> tuple[AgentEvaluatorResult, ...]:
    return tuple(
        AgentEvaluatorResult(
            evaluator=name,
            status=(
                AgentEvaluationStatus.FAILED
                if name is failed
                else AgentEvaluationStatus.SKIPPED
                if name is AgentEvaluatorName.GENERIC_REPAIR_REUSE
                else AgentEvaluationStatus.PASSED
            ),
            reason="deterministic evidence result",
            evidence_path=f"migration/evidence/evaluations/{name.value}.json",
            required=name is not AgentEvaluatorName.GENERIC_REPAIR_REUSE,
        )
        for name in AgentEvaluatorName
    )


def _case(
    *,
    failed: AgentEvaluatorName | None = None,
) -> AgentMigrationEvaluation:
    blockers = (failed.value,) if failed is not None else ()
    return AgentMigrationEvaluation(
        migration_id="invoice-001",
        source_format="xlsm",
        feature_families=("vba", "userforms"),
        model_id="openai:gpt-5.6",
        provider="openai",
        model_version="gpt-5.6",
        team_configuration="lead-specialists-v1",
        evaluators=_evaluators(failed=failed),
        public=_summary(
            "public",
            passed=12 if failed is None else 11,
            failed=0 if failed is None else 1,
            skipped=1,
        ),
        hidden=_summary("hidden", passed=1),
        release_blockers=blockers,
        release_ready=failed is None,
    )


def _report(case: AgentMigrationEvaluation) -> AgentBenchmarkReport:
    public = _summary(
        "public",
        passed=12 if case.release_ready else 11,
        failed=0 if case.release_ready else 1,
        skipped=1,
    )
    hidden = _summary("hidden", passed=1)
    return AgentBenchmarkReport(
        cases=(case,),
        public_by_configuration={"lead-specialists-v1": public},
        hidden_by_configuration={"lead-specialists-v1": hidden},
        public_by_format={"xlsm": public},
        hidden_by_format={"xlsm": hidden},
        public_by_feature_family={"vba": public, "userforms": public},
        hidden_by_feature_family={"vba": hidden, "userforms": hidden},
    )


def test_green_fourteen_evaluator_report_is_release_ready(tmp_path: Path) -> None:
    report = _report(_case())
    path = tmp_path / "agent-evaluation.json"
    path.write_text(report.model_dump_json(indent=2), encoding="utf-8")

    loaded = require_agent_benchmark_release(path)

    assert loaded.target == "libreoffice"
    assert loaded.target_libreoffice_build == "26.2.4.2"
    assert loaded.release_ready is True
    assert len(loaded.cases[0].evaluators) == 14


def test_failed_required_evaluator_blocks_release(tmp_path: Path) -> None:
    report = _report(_case(failed=AgentEvaluatorName.SECURITY_POLICY))
    path = tmp_path / "agent-evaluation.json"
    path.write_text(report.model_dump_json(indent=2), encoding="utf-8")

    with pytest.raises(RuntimeError, match="agent benchmark release gates failed"):
        require_agent_benchmark_release(path)


def test_five_statuses_are_preserved_and_hidden_definitions_are_forbidden() -> None:
    summary = _summary(
        "hidden",
        passed=1,
        failed=2,
        skipped=3,
        unavailable=4,
        not_run=5,
    )

    assert set(summary.counts) == set(AgentEvaluationStatus)
    assert summary.decisive_pass_rate == pytest.approx(1 / 3)
    with pytest.raises(ValidationError):
        AgentPartitionSummary.model_validate(
            {
                **summary.model_dump(mode="json"),
                "hidden_definitions_included": True,
            }
        )


def test_public_and_hidden_groupings_cannot_be_merged() -> None:
    case = _case()
    public = _summary("public", passed=12, skipped=1)

    with pytest.raises(ValidationError, match="wrong partition"):
        AgentBenchmarkReport(
            cases=(case,),
            public_by_configuration={"lead-specialists-v1": public},
            hidden_by_configuration={"lead-specialists-v1": public},
            public_by_format={"xlsm": public},
            hidden_by_format={"xlsm": public},
            public_by_feature_family={"vba": public, "userforms": public},
            hidden_by_feature_family={"vba": public, "userforms": public},
        )
