"""Serious migration episodes, public acceptance contracts, and result reports.

The demo corpus is deliberately evidence-first: a checked-in source and public
scenario describe work to perform, while only an execution result may claim
that LibreOffice behavior was verified.
"""

from __future__ import annotations

import hashlib
import json
from collections import Counter
from pathlib import Path, PurePosixPath
from typing import Literal, cast

import yaml
from pydantic import BaseModel, ConfigDict, Field, model_validator

from xlsliberator.conformance_corpus import EvidenceStatus

DemoFormat = Literal["xls", "xlsx", "xlsm", "xlsb"]
SubsetName = Literal["pr", "nightly", "security"]
TargetStatus = Literal["not_verified", "verified"]


def _safe_relative_path(value: str) -> str:
    path = PurePosixPath(value)
    if path.is_absolute() or ".." in path.parts or not path.parts:
        raise ValueError(f"path must stay inside the repository: {value}")
    return value


class DemoSource(BaseModel):
    """One immutable source workbook."""

    model_config = ConfigDict(extra="forbid")

    path: str
    format: DemoFormat
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    license: str

    @model_validator(mode="after")
    def safe_path_and_extension(self) -> DemoSource:
        """Reject path traversal and source/format mismatches."""
        _safe_relative_path(self.path)
        if PurePosixPath(self.path).suffix.lower() != f".{self.format}":
            raise ValueError("source extension does not match declared format")
        return self


class DemoTarget(BaseModel):
    """Optional target that cannot be called known-good without evidence."""

    model_config = ConfigDict(extra="forbid")

    status: TargetStatus = "not_verified"
    path: str | None = None
    sha256: str | None = Field(default=None, pattern=r"^[0-9a-f]{64}$")
    evidence_path: str | None = None

    @model_validator(mode="after")
    def verified_targets_require_identity_and_evidence(self) -> DemoTarget:
        """Fail closed on unsupported known-good target claims."""
        if self.status == "verified":
            if not self.path or not self.sha256 or not self.evidence_path:
                raise ValueError("verified targets require path, sha256, and evidence_path")
            _safe_relative_path(self.path)
            _safe_relative_path(self.evidence_path)
        elif any((self.path, self.sha256, self.evidence_path)):
            raise ValueError("unverified targets must not publish target artifacts")
        return self


class DemoEpisode(BaseModel):
    """Manifest entry for one complete migration episode."""

    model_config = ConfigDict(extra="forbid")

    episode_id: str = Field(pattern=r"^[a-z0-9][a-z0-9-]+$")
    title: str
    root: str
    source: DemoSource
    task: str
    acceptance: str
    dependencies: str
    restrictions: str
    readme: str
    provenance: str
    license_path: str
    target: DemoTarget = Field(default_factory=DemoTarget)
    expected_features: list[str] = Field(min_length=1)
    tags: list[str] = Field(min_length=1)
    subsets: list[SubsetName] = Field(min_length=1)

    @model_validator(mode="after")
    def paths_and_subsets_are_unambiguous(self) -> DemoEpisode:
        """Ensure all episode paths are safe and PR is a nightly subset."""
        for path in (
            self.root,
            self.source.path,
            self.task,
            self.acceptance,
            self.dependencies,
            self.restrictions,
            self.readme,
            self.provenance,
            self.license_path,
        ):
            _safe_relative_path(path)
        if len(self.subsets) != len(set(self.subsets)):
            raise ValueError("duplicate subset")
        if "pr" in self.subsets and "nightly" not in self.subsets:
            raise ValueError("every PR episode must also run nightly")
        return self


class PublicScenario(BaseModel):
    """Observable behavior in a public, non-hidden acceptance contract."""

    model_config = ConfigDict(extra="forbid")

    scenario_id: str = Field(pattern=r"^[a-z0-9][a-z0-9-]+$")
    title: str
    feature: str
    action: str = Field(min_length=8)
    assertions: list[str] = Field(min_length=2)
    forbidden_effects: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def require_behavioral_assertions(self) -> PublicScenario:
        """Reject existence-only demos that do not exercise user behavior."""
        normalized = " ".join(self.assertions).lower()
        existence_only = all(
            token not in normalized
            for token in (
                "value",
                "state",
                "event",
                "control",
                "row",
                "record",
                "result",
                "output",
                "blocked",
                "denied",
                "reopen",
                "formula",
                "pdf",
                "message",
            )
        )
        if existence_only:
            raise ValueError("scenario assertions must verify behavior, not only a file")
        return self


class PublicAcceptance(BaseModel):
    """Public contract supplied to an implementation agent."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    visibility: Literal["public"] = "public"
    episode_id: str
    target: Literal["libreoffice"] = "libreoffice"
    target_version: Literal["26.2.4.2"] = "26.2.4.2"
    expected_restrictions: list[str] = Field(min_length=1)
    scenarios: list[PublicScenario] = Field(min_length=1)

    @model_validator(mode="after")
    def unique_scenarios(self) -> PublicAcceptance:
        """Reject duplicate public scenario identifiers."""
        ids = [scenario.scenario_id for scenario in self.scenarios]
        if len(ids) != len(set(ids)):
            raise ValueError("duplicate public scenario id")
        return self

    @classmethod
    def load(cls, path: Path) -> PublicAcceptance:
        """Load a YAML acceptance contract."""
        payload = yaml.safe_load(path.read_text(encoding="utf-8"))
        return cls.model_validate(payload)


class DemoCorpusManifest(BaseModel):
    """Machine-readable catalog backing search, subsets, and validation."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    corpus_id: str
    hidden_tests_in_repository: Literal[False] = False
    episodes: list[DemoEpisode] = Field(min_length=8)

    @model_validator(mode="after")
    def unique_episode_ids(self) -> DemoCorpusManifest:
        """Reject ambiguous episodes and require all flagship formats."""
        ids = [episode.episode_id for episode in self.episodes]
        if len(ids) != len(set(ids)):
            raise ValueError("duplicate demo episode id")
        if {episode.source.format for episode in self.episodes} < {
            "xls",
            "xlsx",
            "xlsm",
            "xlsb",
        }:
            raise ValueError("demo corpus must cover XLS, XLSX, XLSM, and XLSB")
        return self

    @classmethod
    def load(cls, path: Path) -> DemoCorpusManifest:
        """Load the episode manifest."""
        return cls.model_validate_json(path.read_text(encoding="utf-8"))

    def search_index(self) -> dict[str, object]:
        """Return stable non-confidential metadata suitable for a corpus MCP."""
        return {
            "schema_version": self.schema_version,
            "corpus_id": self.corpus_id,
            "episodes": [
                {
                    "episode_id": episode.episode_id,
                    "title": episode.title,
                    "format": episode.source.format,
                    "features": sorted(episode.expected_features),
                    "tags": sorted(episode.tags),
                    "subsets": sorted(episode.subsets),
                    "task": episode.task,
                    "acceptance": episode.acceptance,
                }
                for episode in sorted(self.episodes, key=lambda item: item.episode_id)
            ],
        }

    def verify(self, repository_root: Path) -> list[str]:
        """Return complete layout, integrity, and public-contract errors."""
        errors: list[str] = []
        for episode in self.episodes:
            root = repository_root / episode.root
            required_files = (
                episode.task,
                episode.acceptance,
                episode.restrictions,
                episode.readme,
                episode.provenance,
                episode.license_path,
            )
            if not root.is_dir():
                errors.append(f"{episode.episode_id}: episode directory is absent")
            if not (repository_root / episode.dependencies).is_dir():
                errors.append(f"{episode.episode_id}: dependency bundle is absent")
            elif not any(
                item.is_file() for item in (repository_root / episode.dependencies).rglob("*")
            ):
                errors.append(f"{episode.episode_id}: dependency bundle is empty")
            for relative in required_files:
                if not (repository_root / relative).is_file():
                    errors.append(f"{episode.episode_id}: required file is absent: {relative}")

            source = repository_root / episode.source.path
            if not source.is_file():
                errors.append(f"{episode.episode_id}: source workbook is absent")
            elif hashlib.sha256(source.read_bytes()).hexdigest() != episode.source.sha256:
                errors.append(f"{episode.episode_id}: source sha256 mismatch")

            acceptance_path = repository_root / episode.acceptance
            if acceptance_path.is_file():
                try:
                    acceptance = PublicAcceptance.load(acceptance_path)
                except (ValueError, yaml.YAMLError) as exc:
                    errors.append(f"{episode.episode_id}: invalid acceptance: {exc}")
                else:
                    if acceptance.episode_id != episode.episode_id:
                        errors.append(f"{episode.episode_id}: acceptance episode id mismatch")
                    declared = {scenario.feature for scenario in acceptance.scenarios}
                    missing = set(episode.expected_features) - declared
                    if missing:
                        errors.append(
                            f"{episode.episode_id}: features lack public scenarios: "
                            + ", ".join(sorted(missing))
                        )
        return errors


class DemoScenarioResult(BaseModel):
    """One evidence disposition for a public scenario."""

    model_config = ConfigDict(extra="forbid")

    episode_id: str
    scenario_id: str
    feature: str
    source_format: DemoFormat
    target: Literal["libreoffice"] = "libreoffice"
    target_version: Literal["26.2.4.2"] = "26.2.4.2"
    status: EvidenceStatus
    evidence_path: str | None = None
    failure_signature: str | None = None

    @model_validator(mode="after")
    def decisive_results_require_evidence(self) -> DemoScenarioResult:
        """Keep passing and failing claims bound to durable evidence."""
        if self.status in {"passed", "failed"} and not self.evidence_path:
            raise ValueError("passed and failed demo results require evidence_path")
        if self.status == "failed" and not self.failure_signature:
            raise ValueError("failed demo results require failure_signature")
        if self.evidence_path:
            _safe_relative_path(self.evidence_path)
        return self


class FeatureDisposition(BaseModel):
    """Counts and conservative status for one feature or source format."""

    model_config = ConfigDict(extra="forbid")

    counts: dict[EvidenceStatus, int]
    capability_status: Literal["passed", "failed", "unsupported", "unavailable", "not_measured"]


class DemoCorpusReport(BaseModel):
    """Evidence-derived feature report without manual success percentages."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    corpus_id: str
    target: Literal["libreoffice"] = "libreoffice"
    target_version: Literal["26.2.4.2"] = "26.2.4.2"
    result_count: int
    episode_status: dict[str, FeatureDisposition]
    feature_status: dict[str, FeatureDisposition]
    format_status: dict[str, FeatureDisposition]


def _disposition(results: list[DemoScenarioResult]) -> FeatureDisposition:
    counts: dict[EvidenceStatus, int] = dict.fromkeys(
        ("passed", "failed", "skipped", "unavailable", "unsupported", "waived"), 0
    )
    counts.update(Counter(result.status for result in results))
    if not results:
        status: Literal["passed", "failed", "unsupported", "unavailable", "not_measured"] = (
            "not_measured"
        )
    elif counts["failed"]:
        status = "failed"
    elif counts["unsupported"]:
        status = "unsupported"
    elif counts["unavailable"] or counts["skipped"] or counts["waived"]:
        status = "unavailable"
    elif counts["passed"] == len(results):
        status = "passed"
    else:
        status = "not_measured"
    return FeatureDisposition(counts=counts, capability_status=status)


def generate_demo_corpus_report(
    manifest: DemoCorpusManifest,
    results: list[DemoScenarioResult],
) -> DemoCorpusReport:
    """Aggregate results by episode, feature, and source format."""
    known_episodes = {episode.episode_id: episode for episode in manifest.episodes}
    for result in results:
        episode = known_episodes.get(result.episode_id)
        if episode is None:
            raise ValueError(f"unknown episode result: {result.episode_id}")
        if result.source_format != episode.source.format:
            raise ValueError(f"source format mismatch: {result.episode_id}")

    def grouped(keys: list[str], selector: str) -> dict[str, FeatureDisposition]:
        return {
            key: _disposition(
                [
                    result
                    for result in results
                    if (
                        result.episode_id
                        if selector == "episode"
                        else result.feature
                        if selector == "feature"
                        else result.source_format
                    )
                    == key
                ]
            )
            for key in sorted(set(keys))
        }

    return DemoCorpusReport(
        corpus_id=manifest.corpus_id,
        result_count=len(results),
        episode_status=grouped([episode.episode_id for episode in manifest.episodes], "episode"),
        feature_status=grouped(
            [feature for episode in manifest.episodes for feature in episode.expected_features],
            "feature",
        ),
        format_status=grouped([episode.source.format for episode in manifest.episodes], "format"),
    )


def load_demo_results(path: Path) -> list[DemoScenarioResult]:
    """Load versioned demo result dispositions."""
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != "1.0.0":
        raise ValueError("unsupported demo result schema")
    return [DemoScenarioResult.model_validate(item) for item in payload.get("results", [])]


def search_demo_corpus(
    manifest: DemoCorpusManifest,
    *,
    query: str,
    subset: SubsetName | None = None,
) -> list[dict[str, object]]:
    """Search public episode metadata without exposing hidden expectations."""
    terms = {term.casefold() for term in query.split() if term.strip()}
    matches: list[dict[str, object]] = []
    index = manifest.search_index()
    entries = cast(list[dict[str, object]], index["episodes"])
    for item in entries:
        subsets = cast(list[str], item["subsets"])
        features = cast(list[str], item["features"])
        tags = cast(list[str], item["tags"])
        if subset and subset not in subsets:
            continue
        haystack = " ".join(
            [
                str(item["episode_id"]),
                str(item["title"]),
                str(item["format"]),
                " ".join(features),
                " ".join(tags),
            ]
        ).casefold()
        if terms and not all(term in haystack for term in terms):
            continue
        matches.append(item)
    return matches
