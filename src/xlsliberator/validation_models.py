"""Validation and artifact IR models."""

from enum import StrEnum
from typing import Any

from pydantic import BaseModel, Field

from xlsliberator.ir_models import WorkbookIR


class TargetKind(StrEnum):
    """Runtime validation target."""

    LIBREOFFICE = "libreoffice"
    OPENOFFICE = "openoffice"
    BOTH = "both"


class ValidationSeverity(StrEnum):
    """Validation issue severity."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    FATAL = "fatal"


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
    """Formula artifact with source metadata."""

    source_ref: SourceRef
    formula_text: str
    dialect: str = "excel_a1"
    tokens: list[Any] | None = None
    ast: dict[str, Any] | None = None
    dependencies: list[str] = Field(default_factory=list)
    volatility_flags: list[str] = Field(default_factory=list)


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


class WorkbookArtifactIR(BaseModel):
    """Workbook-level artifact inventory."""

    workbook: WorkbookIR
    formulas: list[FormulaIR] = Field(default_factory=list)
    controls: list[ControlIR] = Field(default_factory=list)
    event_bindings: list[EventBindingIR] = Field(default_factory=list)
    source_maps: list[SourceMapIR] = Field(default_factory=list)
    unsupported_artifacts: list[UnsupportedArtifactIR] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)


class ValidationGateResult(BaseModel):
    """Result of one validation gate."""

    gate_name: str
    passed: bool
    severity: ValidationSeverity = ValidationSeverity.INFO
    message: str
    details: dict[str, Any] = Field(default_factory=dict)


class ValidationCertification(BaseModel):
    """Certification result for a transformed workbook."""

    certified: bool = False
    target_profiles: list[str] = Field(default_factory=list)
    gate_results: list[ValidationGateResult] = Field(default_factory=list)
    unsupported_artifacts: list[UnsupportedArtifactIR] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)
