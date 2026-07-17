"""Validation and artifact IR models."""

from __future__ import annotations

from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator

from xlsliberator.ir_models import WorkbookIR


class TargetKind(StrEnum):
    """Runtime validation target."""

    LIBREOFFICE = "libreoffice"


class ValidationSeverity(StrEnum):
    """Validation issue severity."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    FATAL = "fatal"


class GateExecutionStatus(StrEnum):
    """Canonical execution state for a validation gate."""

    PASSED = "passed"
    FAILED = "failed"
    SKIPPED = "skipped"
    UNAVAILABLE = "unavailable"
    NOT_RUN = "not_run"


class CertificationTier(StrEnum):
    """Evidence strength reached by a certification report."""

    STRUCTURAL = "structural"
    TARGET_RUNTIME_VALIDATED = "target_runtime_validated"
    SOURCE_DIFFERENTIAL_VALIDATED = "source_differential_validated"
    LIBREOFFICE_RUNTIME_VALIDATED = "libreoffice_runtime_validated"


class ArtifactCoverage(StrEnum):
    """How completely one source artifact was inspected."""

    SEMANTIC = "semantic"
    RAW = "raw"
    SEMANTIC_AND_RAW = "semantic_and_raw"
    UNPARSED = "unparsed"


class ArtifactDispositionKind(StrEnum):
    """Explicit outcome for one discovered source artifact."""

    PRESERVED = "preserved"
    TRANSLATED = "translated"
    EMULATED = "emulated"
    EXTERNALIZED_DEPENDENCY = "externalized_dependency"
    WAIVED = "waived"
    FAILED = "failed"


class SourceRef(BaseModel):
    """Reference to a source workbook artifact."""

    source_file: str
    sheet: str | None = None
    cell_range: str | None = None
    module: str | None = None
    procedure: str | None = None
    artifact_type: str
    artifact_id: str


class TargetRef(BaseModel):
    """Reference to a target workbook artifact."""

    target_file: str
    sheet: str | None = None
    cell_range: str | None = None
    script_uri: str | None = None
    artifact_type: str
    artifact_id: str


class UnsupportedArtifactIR(BaseModel):
    """Artifact that was detected but is unsupported or unverified."""

    source_ref: SourceRef
    reason: str
    severity: ValidationSeverity = ValidationSeverity.WARNING
    details: dict[str, Any] = Field(default_factory=dict)


class FormulaIR(BaseModel):
    """Versioned semantic formula artifact with source and target context."""

    schema_version: Literal["2.0.0"] = "2.0.0"
    source_ref: SourceRef
    target_ref: TargetRef | None = None
    formula_text: str
    original_formula_text: str | None = None
    target_formula_text: str | None = None
    dialect: str = "excel_a1"
    tokens: list[Any] | None = None
    ast: dict[str, Any] | None = None
    sheet_context: str | None = None
    cell_context: str | None = None
    name_context: str | None = None
    dependencies: list[str] = Field(default_factory=list)
    volatility_flags: list[str] = Field(default_factory=list)
    semantic_features: list[str] = Field(default_factory=list)
    semantic_diagnostics: list[str] = Field(default_factory=list)
    runtime_evidence_requirements: list[str] = Field(default_factory=list)
    array_metadata: dict[str, Any] = Field(default_factory=dict)
    calculation_settings: dict[str, Any] = Field(default_factory=dict)
    calculation_order: dict[str, Any] = Field(default_factory=dict)
    unsupported_reasons: list[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def fill_original_and_context(self) -> FormulaIR:
        if self.original_formula_text is None:
            self.original_formula_text = self.formula_text
        if self.sheet_context is None:
            self.sheet_context = self.source_ref.sheet
        if self.cell_context is None:
            self.cell_context = self.source_ref.cell_range
        return self


class ControlIR(BaseModel):
    """Spreadsheet form control artifact."""

    id: str
    name: str
    control_type: str
    sheet: str | None = None
    anchor: str | None = None
    properties: dict[str, Any] = Field(default_factory=dict)
    linked_cell: str | None = None
    list_fill_range: str | None = None
    source_ref: SourceRef | None = None
    target_ref: TargetRef | None = None


class EventBindingIR(BaseModel):
    """Source-to-target event binding artifact."""

    id: str
    source_ref: SourceRef
    event_name: str
    source_handler: str
    target_script_uri: str | None = None
    control_id: str | None = None
    details: dict[str, Any] = Field(default_factory=dict)


class SourceMapIR(BaseModel):
    """Stable source map entry for translated artifacts."""

    source_ref: SourceRef
    target_ref: TargetRef | None = None
    marker: str
    metadata: dict[str, Any] = Field(default_factory=dict)


class CanonicalArtifactIR(BaseModel):
    """One stable semantic or raw workbook artifact."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    artifact_id: str
    family: str
    artifact_type: str
    locator: str
    coverage: ArtifactCoverage
    source_ref: SourceRef
    parent_artifact_id: str | None = None
    semantic_data: dict[str, Any] = Field(default_factory=dict)
    raw_path: str | None = None
    raw_sha256: str | None = None
    raw_size: int | None = None
    relationship_ids: list[str] = Field(default_factory=list)
    known: bool = True


class ArtifactDisposition(BaseModel):
    """Loss-accounting decision for one discovered source artifact."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    source_artifact_id: str
    disposition: ArtifactDispositionKind
    target_refs: list[TargetRef] = Field(default_factory=list)
    evidence_references: list[str] = Field(default_factory=list)
    reason: str | None = None


class InventoryDiff(BaseModel):
    """Deterministic source-versus-target artifact comparison."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    source_inventory_sha256: str
    target_inventory_sha256: str
    matched: dict[str, str] = Field(default_factory=dict)
    missing_source_artifact_ids: list[str] = Field(default_factory=list)
    added_target_artifact_ids: list[str] = Field(default_factory=list)
    dispositions: list[ArtifactDisposition] = Field(default_factory=list)


class WorkbookArtifactIR(BaseModel):
    """Workbook-level artifact inventory."""

    schema_version: Literal["3.0.0"] = "3.0.0"
    inventory_role: Literal["source", "target"] = "source"
    source_sha256: str | None = None
    workbook: WorkbookIR
    formulas: list[FormulaIR] = Field(default_factory=list)
    controls: list[ControlIR] = Field(default_factory=list)
    event_bindings: list[EventBindingIR] = Field(default_factory=list)
    source_maps: list[SourceMapIR] = Field(default_factory=list)
    unsupported_artifacts: list[UnsupportedArtifactIR] = Field(default_factory=list)
    artifacts: list[CanonicalArtifactIR] = Field(default_factory=list)
    dispositions: list[ArtifactDisposition] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def artifact_ids_are_unique(self) -> WorkbookArtifactIR:
        artifact_ids = [artifact.artifact_id for artifact in self.artifacts]
        if len(artifact_ids) != len(set(artifact_ids)):
            raise ValueError("canonical artifact IDs must be unique")
        return self


class ValidationGateResult(BaseModel):
    """Result of one validation gate."""

    gate_name: str
    status: GateExecutionStatus = GateExecutionStatus.NOT_RUN
    passed: bool = False
    required: bool = True
    severity: ValidationSeverity = ValidationSeverity.INFO
    message: str
    details: dict[str, Any] = Field(default_factory=dict)
    evidence_references: list[str] = Field(default_factory=list)

    @model_validator(mode="before")
    @classmethod
    def project_legacy_passed_field(cls, value: Any) -> Any:
        """Accept legacy ``passed`` input while keeping status canonical."""
        if not isinstance(value, dict):
            return value
        normalized = dict(value)
        if "status" not in normalized:
            if "passed" in normalized:
                normalized["status"] = (
                    GateExecutionStatus.PASSED
                    if normalized["passed"]
                    else GateExecutionStatus.FAILED
                )
            else:
                normalized["status"] = GateExecutionStatus.NOT_RUN
        status = GateExecutionStatus(normalized["status"])
        normalized["passed"] = status == GateExecutionStatus.PASSED
        return normalized


class RepairProvenance(BaseModel):
    """Reference to a repair run; it carries provenance, never authority."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    agent_run_id: str
    candidate_patch_id: str
    agent_run_reference: str
    accepted_patch_sha256: str
    deterministic_gate_names: list[str] = Field(default_factory=list)


class ValidationCertification(BaseModel):
    """Certification result for a transformed workbook."""

    schema_version: str = "2.0.0"
    certified: bool = False
    tier: CertificationTier = CertificationTier.STRUCTURAL
    target_profiles: list[str] = Field(default_factory=list)
    gate_results: list[ValidationGateResult] = Field(default_factory=list)
    repair_provenance: list[RepairProvenance] = Field(default_factory=list)
    unsupported_artifacts: list[UnsupportedArtifactIR] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def enforce_gate_semantics(self) -> ValidationCertification:
        """Derive certification from required gate statuses, never caller intent."""
        required_gates = [gate for gate in self.gate_results if gate.required]
        self.certified = bool(required_gates) and all(
            gate.status == GateExecutionStatus.PASSED for gate in required_gates
        )
        return self
