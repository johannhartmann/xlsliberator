"""Validate Open-SWE migration benchmark evidence before release."""

from __future__ import annotations

import json
from enum import StrEnum
from pathlib import Path
from typing import Literal, Self

from pydantic import BaseModel, ConfigDict, Field, model_validator


class AgentEvaluationStatus(StrEnum):
    PASSED = "passed"
    FAILED = "failed"
    SKIPPED = "skipped"
    UNAVAILABLE = "unavailable"
    NOT_RUN = "not_run"


class AgentEvaluatorName(StrEnum):
    SPECIALIST_DELEGATION = "correct-specialist-delegation"
    SKILL_SELECTION = "relevant-skill-selection"
    NO_FAKE_SUCCESS = "no-fake-success"
    NO_TEST_WEAKENING = "no-test-weakening"
    SOURCE_DERIVED_TEST_QUALITY = "source-derived-test-quality"
    HIDDEN_ACCEPTANCE = "hidden-acceptance-pass"
    MUTATION_KILL_RATE = "mutation-kill-rate"
    SAVE_REOPEN = "save-reopen-pass"
    PROPRIETARY_DEPENDENCY_REMOVAL = "proprietary-dependency-removal"
    REVIEWER_AGREEMENT = "reviewer-agreement"
    GENERIC_REPAIR_REUSE = "generic-repair-reuse"
    MANUAL_INTERVENTION_RATE = "manual-intervention-rate"
    COST_LATENCY = "cost-latency-per-success"
    SECURITY_POLICY = "security-policy-adherence"


class AgentEvaluatorResult(BaseModel):
    """One exported deterministic evaluator result."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    evaluator: AgentEvaluatorName
    status: AgentEvaluationStatus
    reason: str = Field(min_length=1, max_length=1000)
    evidence_path: str = Field(pattern=r"^migration/evidence/[a-z0-9][a-z0-9._/-]{0,255}$")
    required: bool = True
    score: float | None = Field(default=None, ge=0.0, le=1.0)

    @model_validator(mode="after")
    def evidence_path_is_confined(self) -> Self:
        if ".." in self.evidence_path.split("/"):
            raise ValueError("agent evaluation evidence path cannot traverse")
        return self


class AgentPartitionSummary(BaseModel):
    """Status counts for one benchmark partition."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    partition: Literal["public", "hidden"]
    counts: dict[AgentEvaluationStatus, int]
    decisive_pass_rate: float | None = Field(default=None, ge=0.0, le=1.0)
    hidden_definitions_included: Literal[False] = False

    @model_validator(mode="after")
    def all_statuses_are_explicit(self) -> Self:
        if set(self.counts) != set(AgentEvaluationStatus):
            raise ValueError("partition summary must retain all five statuses")
        if any(count < 0 for count in self.counts.values()):
            raise ValueError("partition status counts cannot be negative")
        decisive = (
            self.counts[AgentEvaluationStatus.PASSED] + self.counts[AgentEvaluationStatus.FAILED]
        )
        expected = self.counts[AgentEvaluationStatus.PASSED] / decisive if decisive else None
        if self.decisive_pass_rate != expected:
            raise ValueError("decisive pass rate must derive from passed and failed only")
        return self


class AgentMigrationEvaluation(BaseModel):
    """One migration case exported by the independent benchmark harness."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    schema_version: Literal["1.0.0"] = "1.0.0"
    migration_id: str = Field(min_length=1, max_length=128)
    source_format: str = Field(min_length=1, max_length=32)
    feature_families: tuple[str, ...] = Field(min_length=1)
    target: Literal["libreoffice"] = "libreoffice"
    target_libreoffice_build: Literal["26.2.4.2"] = "26.2.4.2"
    model_id: str = Field(min_length=1, max_length=200)
    provider: str = Field(min_length=1, max_length=100)
    model_version: str = Field(min_length=1, max_length=200)
    team_configuration: str = Field(min_length=1, max_length=100)
    evaluators: tuple[AgentEvaluatorResult, ...] = Field(
        min_length=14,
        max_length=14,
    )
    public: AgentPartitionSummary
    hidden: AgentPartitionSummary
    release_blockers: tuple[str, ...]
    release_ready: bool

    @model_validator(mode="after")
    def release_is_fail_closed(self) -> Self:
        actual = {result.evaluator for result in self.evaluators}
        if actual != set(AgentEvaluatorName) or len(actual) != len(self.evaluators):
            raise ValueError("all fourteen agent evaluators are required exactly once")
        if self.public.partition != "public" or self.hidden.partition != "hidden":
            raise ValueError("public and hidden benchmark partitions cannot be merged")
        required_failures = [
            result.evaluator.value
            for result in self.evaluators
            if result.required and result.status is not AgentEvaluationStatus.PASSED
        ]
        expected_blocked = bool(required_failures or self.release_blockers)
        if self.release_ready == expected_blocked:
            raise ValueError("release_ready contradicts evaluator or release blockers")
        return self


class AgentBenchmarkReport(BaseModel):
    """Cross-repository release artifact generated by the Open-SWE benchmark."""

    model_config = ConfigDict(extra="forbid", frozen=True)

    schema_version: Literal["1.0.0"] = "1.0.0"
    target: Literal["libreoffice"] = "libreoffice"
    target_libreoffice_build: Literal["26.2.4.2"] = "26.2.4.2"
    cases: tuple[AgentMigrationEvaluation, ...] = Field(min_length=1)
    public_by_configuration: dict[str, AgentPartitionSummary]
    hidden_by_configuration: dict[str, AgentPartitionSummary]
    public_by_format: dict[str, AgentPartitionSummary]
    hidden_by_format: dict[str, AgentPartitionSummary]
    public_by_feature_family: dict[str, AgentPartitionSummary]
    hidden_by_feature_family: dict[str, AgentPartitionSummary]

    @model_validator(mode="after")
    def groupings_and_release_gates_are_complete(self) -> Self:
        configurations = {case.team_configuration for case in self.cases}
        source_formats = {case.source_format for case in self.cases}
        feature_families = {family for case in self.cases for family in case.feature_families}
        for expected, public, hidden, label in (
            (
                configurations,
                self.public_by_configuration,
                self.hidden_by_configuration,
                "configuration",
            ),
            (source_formats, self.public_by_format, self.hidden_by_format, "format"),
            (
                feature_families,
                self.public_by_feature_family,
                self.hidden_by_feature_family,
                "feature family",
            ),
        ):
            if set(public) != expected or set(hidden) != expected:
                raise ValueError(f"public/hidden {label} groups must cover every case")
            if any(summary.partition != "public" for summary in public.values()):
                raise ValueError(f"public {label} groups have the wrong partition")
            if any(summary.partition != "hidden" for summary in hidden.values()):
                raise ValueError(f"hidden {label} groups have the wrong partition")
        return self

    @property
    def release_ready(self) -> bool:
        """Return true only if every exported migration is release-ready."""
        return bool(self.cases) and all(case.release_ready for case in self.cases)


def load_agent_benchmark_report(path: Path) -> AgentBenchmarkReport:
    """Load and validate a privacy-safe Open-SWE benchmark artifact."""
    return AgentBenchmarkReport.model_validate(json.loads(path.read_text(encoding="utf-8")))


def require_agent_benchmark_release(path: Path) -> AgentBenchmarkReport:
    """Reject missing, malformed, stale-target, or non-green agent evidence."""
    if not path.is_file():
        raise RuntimeError(f"required agent benchmark report is absent: {path}")
    report = load_agent_benchmark_report(path)
    if not report.release_ready:
        blockers = {
            case.migration_id: list(case.release_blockers)
            for case in report.cases
            if not case.release_ready
        }
        raise RuntimeError(f"agent benchmark release gates failed: {blockers}")
    return report
