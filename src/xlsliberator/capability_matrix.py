"""Generate capability claims and release gates exclusively from evidence."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from xlsliberator.conformance_corpus import CorpusManifest, EvidenceStatus
from xlsliberator.formula_corpus import FormulaCorpusStatistics

CoverageStatus = Literal["passed", "failed", "skipped", "unavailable", "unsupported", "waived"]
CertificationTier = Literal[
    "structural-inventory",
    "target-runtime-validated",
    "source-differential-validated",
    "libreoffice-runtime-validated",
]
CERTIFICATION_TIERS: tuple[CertificationTier, ...] = (
    "structural-inventory",
    "target-runtime-validated",
    "source-differential-validated",
    "libreoffice-runtime-validated",
)


class RuntimeEvidenceIdentity(BaseModel):
    """Exact immutable identity required for passing target evidence."""

    model_config = ConfigDict(extra="forbid")

    image_reference: str
    image_digest: str = Field(pattern=r"^sha256:[0-9a-f]{64}$")
    base_image_digest: str = Field(pattern=r"^sha256:[0-9a-f]{64}$")
    architecture: str
    python_version: str
    pyuno_identity: dict[str, str]
    office_binary_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    package_set: list[str]
    runtime_variant: str
    source_commit: str
    patch_set_sha256: str

    @model_validator(mode="after")
    def require_source_and_pyuno_identity(self) -> RuntimeEvidenceIdentity:
        """Reject partial identities that cannot bind evidence to one runtime."""
        if self.source_commit != "official-binary-distribution" and (
            len(self.source_commit) != 40
            or any(character not in "0123456789abcdef" for character in self.source_commit)
        ):
            raise ValueError("source_commit must be an exact commit or official distribution")
        if self.patch_set_sha256 != "none" and (
            len(self.patch_set_sha256) != 64
            or any(character not in "0123456789abcdef" for character in self.patch_set_sha256)
        ):
            raise ValueError("patch_set_sha256 must be none or a sha256")
        required_pyuno = {"uno_module_sha256", "pyuno_native_sha256"}
        if not required_pyuno.issubset(self.pyuno_identity):
            raise ValueError("runtime identity requires UNO and native PyUNO hashes")
        if not self.package_set:
            raise ValueError("runtime identity requires a non-empty package set")
        return self


class CapabilityMeasurement(BaseModel):
    """One evidence-derived capability cell."""

    model_config = ConfigDict(extra="forbid")

    evidence_id: str
    fixture_id: str
    source_format: str
    artifact_family: str
    scenario: str
    environment: str
    target: Literal["libreoffice"] = "libreoffice"
    target_version: Literal["26.2.4.2"] = "26.2.4.2"
    runtime: RuntimeEvidenceIdentity | None = None
    parse_coverage: CoverageStatus
    output_coverage: CoverageStatus
    target_runtime: CoverageStatus
    source_differential: CoverageStatus
    evidence_bundle: str | None = None
    waiver: str | None = None

    @model_validator(mode="after")
    def decisive_results_require_evidence(self) -> CapabilityMeasurement:
        """Prevent a passing or failing claim without a runtime and bundle."""
        decisive = {
            self.parse_coverage,
            self.output_coverage,
            self.target_runtime,
            self.source_differential,
        } & {"passed", "failed"}
        if decisive and not self.evidence_bundle:
            raise ValueError("passed/failed capability measurements require an evidence bundle")
        if self.target_runtime == "passed" and self.runtime is None:
            raise ValueError("target-runtime passes require an exact runtime identity")
        if (
            "waived"
            in {
                self.parse_coverage,
                self.output_coverage,
                self.target_runtime,
                self.source_differential,
            }
            and not self.waiver
        ):
            raise ValueError("waived measurements require an explicit waiver")
        return self

    @property
    def tiers(self) -> list[CertificationTier]:
        """Derive tiers; higher tiers never imply unavailable source evidence."""
        tiers: list[CertificationTier] = []
        if self.parse_coverage == "passed":
            tiers.append("structural-inventory")
        if self.target_runtime == "passed":
            tiers.append("target-runtime-validated")
            if self.runtime and self.target == "libreoffice":
                tiers.append("libreoffice-runtime-validated")
        if self.source_differential == "passed":
            tiers.append("source-differential-validated")
        return tiers


class ReleaseInputs(BaseModel):
    """Results of independent blocking release gates."""

    model_config = ConfigDict(extra="forbid")

    p0_tests_passed: bool
    fail_open_paths_absent: bool
    source_artifacts_accounted: bool
    evidence_schemas_valid: bool
    security_suite_passed: bool
    agent_evaluation_passed: bool = False


class ReleaseGate(BaseModel):
    """One explicit gate result."""

    model_config = ConfigDict(extra="forbid")

    name: str
    passed: bool
    reason: str


class CoverageSummary(BaseModel):
    """Separate disposition counts and a rate over decisive results only."""

    model_config = ConfigDict(extra="forbid")

    counts: dict[EvidenceStatus, int]
    decisive_pass_rate: float | None


class CapabilityReport(BaseModel):
    """Versioned generated report used for docs and release decisions."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    report_id: str = "current"
    generated_from_evidence: Literal[True] = True
    target: Literal["libreoffice"] = "libreoffice"
    target_version: Literal["26.2.4.2"] = "26.2.4.2"
    measurements: list[CapabilityMeasurement]
    formula_corpus: FormulaCorpusStatistics = Field(
        default_factory=lambda: FormulaCorpusStatistics(corpus_path="not-recorded")
    )
    summaries: dict[str, CoverageSummary]
    tier_counts: dict[CertificationTier, int]
    release_gates: list[ReleaseGate]
    release_ready: bool
    previous_release: str | None = None
    trend: dict[str, int] = Field(default_factory=dict)

    def to_markdown(self) -> str:
        """Render a stable document without manually maintained percentages."""
        lines = [
            "# XLSLiberator Capability Matrix",
            "",
            "> Generated from corpus and evidence data. Do not edit by hand.",
            "",
            f"Target: LibreOffice `{self.target_version}` in pinned Docker images.",
            "",
            "## Status semantics",
            "",
            "`unavailable`, `skipped`, `unsupported`, `waived`, and `failed` are distinct. "
            "Rates use only decisive `passed` and `failed` results.",
            "",
            "## Measurements",
            "",
            "| Format | Artifact | Scenario | Environment | Parse | Output | Target runtime | Source differential | Tiers |",
            "|---|---|---|---|---|---|---|---|---|",
        ]
        for item in self.measurements:
            lines.append(
                "| "
                + " | ".join(
                    [
                        item.source_format.upper(),
                        item.artifact_family,
                        item.scenario,
                        item.environment,
                        item.parse_coverage,
                        item.output_coverage,
                        item.target_runtime,
                        item.source_differential,
                        ", ".join(item.tiers) or "none",
                    ]
                )
                + " |"
            )
        lines.extend(["", "## Evidence identities", ""])
        for item in self.measurements:
            if item.runtime is None:
                lines.append(f"- `{item.evidence_id}`: runtime identity unavailable")
            else:
                lines.append(
                    f"- `{item.evidence_id}`: `{item.runtime.image_digest}`; "
                    f"base `{item.runtime.base_image_digest}`; {item.runtime.architecture}; "
                    f"Python {item.runtime.python_version}; "
                    f"variant `{item.runtime.runtime_variant}`; "
                    f"office `{item.runtime.office_binary_sha256}`; "
                    f"UNO `{item.runtime.pyuno_identity['uno_module_sha256']}`; "
                    f"PyUNO `{item.runtime.pyuno_identity['pyuno_native_sha256']}`; "
                    f"source `{item.runtime.source_commit}`; patch "
                    f"`{item.runtime.patch_set_sha256}`; packages "
                    f"{len(item.runtime.package_set)}"
                )
        lines.extend(
            [
                "",
                "## Formula corpus",
                "",
                f"- Minimized regression fixtures: {self.formula_corpus.minimized_regression_fixtures}",
                f"- Registered rules: {self.formula_corpus.registered_rules}",
                f"- Covered rules: {self.formula_corpus.covered_rules}",
                f"- Source differential: `{self.formula_corpus.source_differential_status}`",
            ]
        )
        lines.extend(["", "## Release gates", ""])
        lines.extend(
            f"- {'PASS' if gate.passed else 'FAIL'} `{gate.name}`: {gate.reason}"
            for gate in self.release_gates
        )
        lines.extend(
            [
                "",
                f"Release ready: **{'YES' if self.release_ready else 'NO'}**",
                "",
                "Certification tiers:",
                "",
                "- `structural-inventory`: source artifacts were inventoried.",
                "- `target-runtime-validated`: required target scenario passed.",
                "- `source-differential-validated`: target matched a source trace.",
                "- `libreoffice-runtime-validated`: target pass used exact pinned LibreOffice Docker evidence.",
                "",
            ]
        )
        return "\n".join(lines)

    def to_release_notes(self) -> str:
        """Render evidence-derived release readiness without marketing inference."""
        lines = [
            "# Release readiness",
            "",
            "> Generated from the capability report. Do not edit by hand.",
            "",
            f"LibreOffice target: `{self.target_version}`.",
            f"Release ready: **{'YES' if self.release_ready else 'NO'}**.",
            "",
            "## Certification counts",
            "",
        ]
        lines.extend(f"- `{tier}`: {count}" for tier, count in self.tier_counts.items())
        lines.extend(["", "## Blocking gates", ""])
        lines.extend(
            f"- {'PASS' if gate.passed else 'FAIL'} `{gate.name}`: {gate.reason}"
            for gate in self.release_gates
        )
        if self.previous_release:
            lines.extend(["", f"Compared with: `{self.previous_release}`.", ""])
            lines.extend(f"- `{tier}`: {delta:+d}" for tier, delta in self.trend.items())
        lines.append("")
        return "\n".join(lines)


def _coverage_summary(values: list[CoverageStatus]) -> CoverageSummary:
    counts: dict[EvidenceStatus, int] = dict.fromkeys(
        ("passed", "failed", "skipped", "unavailable", "unsupported", "waived"), 0
    )
    counts.update(Counter(values))
    decisive = counts["passed"] + counts["failed"]
    return CoverageSummary(
        counts=counts,
        decisive_pass_rate=(counts["passed"] / decisive if decisive else None),
    )


def generate_capability_report(
    *,
    corpus: CorpusManifest,
    measurements: list[CapabilityMeasurement],
    release_inputs: ReleaseInputs,
    formula_corpus: FormulaCorpusStatistics | None = None,
    previous: CapabilityReport | None = None,
    report_id: str = "current",
) -> CapabilityReport:
    """Build the only publishable report and fail closed on missing corpus evidence."""
    fixture_ids = {fixture.fixture_id for fixture in corpus.fixtures}
    unknown = sorted({item.fixture_id for item in measurements} - fixture_ids)
    if unknown:
        raise ValueError(f"capability evidence references unknown fixtures: {unknown}")
    blocking = {fixture.fixture_id for fixture in corpus.fixtures if fixture.blocking}
    by_fixture = {item.fixture_id: item for item in measurements}
    non_green_blocking = sorted(
        fixture_id
        for fixture_id in blocking
        if fixture_id not in by_fixture or by_fixture[fixture_id].target_runtime != "passed"
    )
    blocking_green = not non_green_blocking
    identities_recorded = all(
        item.runtime is not None for item in measurements if item.target_runtime == "passed"
    )
    gates = [
        ReleaseGate(
            name="p0-tests", passed=release_inputs.p0_tests_passed, reason="P0 suite result"
        ),
        ReleaseGate(
            name="fail-closed-certification",
            passed=release_inputs.fail_open_paths_absent,
            reason="no fail-open certification path",
        ),
        ReleaseGate(
            name="required-corpus",
            passed=blocking_green,
            reason=(
                "all blocking fixtures are accounted green"
                if blocking_green
                else f"missing or non-green: {non_green_blocking}"
            ),
        ),
        ReleaseGate(
            name="source-artifact-accounting",
            passed=release_inputs.source_artifacts_accounted,
            reason="certified fixtures have complete dispositions",
        ),
        ReleaseGate(
            name="evidence-schemas",
            passed=release_inputs.evidence_schemas_valid,
            reason="all evidence models validated",
        ),
        ReleaseGate(
            name="security-suite",
            passed=release_inputs.security_suite_passed,
            reason="blocking security suite result",
        ),
        ReleaseGate(
            name="agent-evaluation",
            passed=release_inputs.agent_evaluation_passed,
            reason="required corpus, security, hidden, and reviewer agent gates",
        ),
        ReleaseGate(
            name="runtime-identities",
            passed=identities_recorded,
            reason="every target pass records binary/source identities",
        ),
    ]
    summaries = {
        field: _coverage_summary([getattr(item, field) for item in measurements])
        for field in (
            "parse_coverage",
            "output_coverage",
            "target_runtime",
            "source_differential",
        )
    }
    tier_counter: Counter[CertificationTier] = Counter(
        tier for measurement in measurements for tier in measurement.tiers
    )
    tier_counts: dict[CertificationTier, int] = {
        tier: tier_counter[tier] for tier in CERTIFICATION_TIERS
    }
    trend: dict[str, int] = {}
    previous_release = None
    if previous is not None:
        previous_release = previous.report_id
        trend = {
            tier: tier_counts[tier] - previous.tier_counts.get(tier, 0) for tier in tier_counts
        }
    return CapabilityReport(
        report_id=report_id,
        measurements=measurements,
        formula_corpus=formula_corpus or FormulaCorpusStatistics(corpus_path="not-recorded"),
        summaries=summaries,
        tier_counts=tier_counts,
        release_gates=gates,
        release_ready=all(gate.passed for gate in gates),
        previous_release=previous_release,
        trend=trend,
    )


def load_measurements(path: Path) -> list[CapabilityMeasurement]:
    """Validate a versioned measurement list."""
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0.0":
        raise ValueError("unsupported capability evidence schema")
    return [CapabilityMeasurement.model_validate(item) for item in payload.get("measurements", [])]
