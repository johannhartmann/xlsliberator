"""Evidence-bound promotion records for reusable workbook migration repairs."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path, PurePosixPath
from typing import Literal, cast

from pydantic import BaseModel, ConfigDict, Field, model_validator

RepairClassification = Literal[
    "workbook-specific",
    "xlsliberator-tool",
    "domain-skill",
    "open-service-adapter",
    "libreoffice",
    "test-validation",
]
RepairStageName = Literal[
    "reproduce",
    "minimize",
    "regression",
    "patch",
    "exact-scenario",
    "affected-corpus",
    "independent-review",
    "upstream-review",
]


def _safe_relative_path(value: str) -> str:
    path = PurePosixPath(value)
    if path.is_absolute() or ".." in path.parts or not path.parts:
        raise ValueError(f"path must stay inside the repository: {value}")
    return value


class RepairArtifact(BaseModel):
    """One immutable file used by the repair promotion workflow."""

    model_config = ConfigDict(extra="forbid")

    path: str
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")

    @model_validator(mode="after")
    def path_is_safe(self) -> RepairArtifact:
        _safe_relative_path(self.path)
        return self


class RepairStage(BaseModel):
    """A decisive workflow stage with durable evidence."""

    model_config = ConfigDict(extra="forbid")

    name: RepairStageName
    status: Literal["passed"]
    evidence_path: str

    @model_validator(mode="after")
    def evidence_path_is_safe(self) -> RepairStage:
        _safe_relative_path(self.evidence_path)
        return self


class LibreOfficeRepairIdentity(BaseModel):
    """Pinned source, patch, runtime, and binary identity for an LO repair."""

    model_config = ConfigDict(extra="forbid")

    full_build: Literal["26.2.4.2"] = "26.2.4.2"
    source_commit: str = Field(pattern=r"^[0-9a-f]{40}$")
    source_archive_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    patch_set_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    stock_runtime_digest: str = Field(pattern=r"^sha256:[0-9a-f]{64}$")
    patched_runtime_digest: str = Field(pattern=r"^sha256:[0-9a-f]{64}$")
    calc_core_stock_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    calc_core_patched_sha256: str = Field(pattern=r"^[0-9a-f]{64}$")


class RepairRecord(BaseModel):
    """Complete reusable-repair record; partial narratives cannot validate."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    repair_id: str = Field(pattern=r"^[a-z0-9][a-z0-9-]+$")
    title: str
    classification: RepairClassification
    failure_signature: str
    fixed_layer: RepairClassification
    source_fixture: RepairArtifact
    minimized_fixture: RepairArtifact
    regression_test: RepairArtifact
    patch: RepairArtifact
    skill_update: RepairArtifact
    exact_scenario_evidence: RepairArtifact
    affected_corpus_evidence: RepairArtifact
    reviewer_evidence: RepairArtifact
    libreoffice: LibreOfficeRepairIdentity | None = None
    upstream_review: str
    stages: list[RepairStage] = Field(min_length=8, max_length=8)

    @model_validator(mode="after")
    def complete_workflow_and_layer_identity(self) -> RepairRecord:
        expected = {
            "reproduce",
            "minimize",
            "regression",
            "patch",
            "exact-scenario",
            "affected-corpus",
            "independent-review",
            "upstream-review",
        }
        names = {stage.name for stage in self.stages}
        if names != expected or len(names) != len(self.stages):
            raise ValueError("repair record must contain every workflow stage exactly once")
        if self.classification != self.fixed_layer:
            raise ValueError("classification and patched layer must agree")
        if self.classification == "libreoffice" and self.libreoffice is None:
            raise ValueError("LibreOffice repairs require pinned source and runtime identity")
        if not self.upstream_review.startswith("https://"):
            raise ValueError("repair requires an HTTPS upstream review")
        return self

    @classmethod
    def load(cls, path: Path) -> RepairRecord:
        """Load one checked-in repair record."""
        return cls.model_validate_json(path.read_text(encoding="utf-8"))

    def verify(self, repository_root: Path) -> list[str]:
        """Verify artifact hashes and the fail-before/pass-after evidence chain."""
        errors: list[str] = []
        artifacts = (
            self.source_fixture,
            self.minimized_fixture,
            self.regression_test,
            self.patch,
            self.skill_update,
            self.exact_scenario_evidence,
            self.affected_corpus_evidence,
            self.reviewer_evidence,
        )
        for artifact in artifacts:
            path = repository_root / artifact.path
            if not path.is_file():
                errors.append(f"{artifact.path}: artifact is absent")
            elif hashlib.sha256(path.read_bytes()).hexdigest() != artifact.sha256:
                errors.append(f"{artifact.path}: sha256 mismatch")
        for stage in self.stages:
            if not (repository_root / stage.evidence_path).is_file():
                errors.append(f"{stage.name}: stage evidence is absent")
        if errors:
            return errors

        evidence = _load_object(repository_root / self.exact_scenario_evidence.path)
        stock = cast(dict[str, object], evidence.get("stock", {}))
        patched = cast(dict[str, object], evidence.get("patched", {}))
        if stock.get("disposition") != "failed-as-expected":
            errors.append("exact scenario does not prove stock failure")
        if patched.get("disposition") != "passed":
            errors.append("exact scenario does not prove patched success")
        if stock.get("source_build_id") != patched.get("source_build_id"):
            errors.append("stock and patched runs use different source commits")
        if self.libreoffice is not None:
            identity = self.libreoffice
            expected = {
                "stock.source_build_id": (stock.get("source_build_id"), identity.source_commit),
                "stock.source_archive_sha256": (
                    stock.get("source_archive_sha256"),
                    identity.source_archive_sha256,
                ),
                "stock.runtime_image_digest": (
                    stock.get("runtime_image_digest"),
                    identity.stock_runtime_digest,
                ),
                "stock.calc_core_sha256": (
                    stock.get("calc_core_sha256"),
                    identity.calc_core_stock_sha256,
                ),
                "patched.runtime_image_digest": (
                    patched.get("runtime_image_digest"),
                    identity.patched_runtime_digest,
                ),
                "patched.calc_core_sha256": (
                    patched.get("calc_core_sha256"),
                    identity.calc_core_patched_sha256,
                ),
                "patched.patch_set_sha256": (
                    patched.get("patch_set_sha256"),
                    identity.patch_set_sha256,
                ),
            }
            errors.extend(
                f"{name} does not match the repair identity"
                for name, (actual, wanted) in expected.items()
                if actual != wanted
            )

        reviewer = _load_object(repository_root / self.reviewer_evidence.path)
        if reviewer.get("repair_id") != self.repair_id or reviewer.get("verdict") != "APPROVE":
            errors.append("independent reviewer evidence does not approve this repair")
        if reviewer.get("reviewer") == reviewer.get("implementation_owner"):
            errors.append("repair reviewer is not independent")

        affected = _load_object(repository_root / self.affected_corpus_evidence.path)
        case_id = _load_object(repository_root / self.minimized_fixture.path).get("case_id")
        if not isinstance(case_id, str) or affected.get("case_id") != case_id:
            errors.append("affected corpus evidence does not include the minimized case")
        source_report = affected.get("source_report")
        source_report_sha256 = affected.get("source_report_sha256")
        if not isinstance(source_report, str) or not isinstance(source_report_sha256, str):
            errors.append("affected corpus evidence lacks source-report identity")
        else:
            try:
                source_report_path = repository_root / _safe_relative_path(source_report)
            except ValueError as exc:
                errors.append(str(exc))
            else:
                if (
                    not source_report_path.is_file()
                    or hashlib.sha256(source_report_path.read_bytes()).hexdigest()
                    != source_report_sha256
                ):
                    errors.append("affected corpus source-report identity does not match")
        return errors


def load_repair_records(records_root: Path) -> list[RepairRecord]:
    """Load every public repair record in stable order."""
    return [
        RepairRecord.load(path)
        for path in sorted(records_root.glob("*/record.json"))
    ]


def _load_object(path: Path) -> dict[str, object]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError(f"expected a JSON object: {path}")
    return cast(dict[str, object], payload)
