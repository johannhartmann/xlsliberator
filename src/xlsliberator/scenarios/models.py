"""Stable Pydantic schemas for scenario execution and evidence."""

from __future__ import annotations

from datetime import datetime
from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from xlsliberator.validation_models import GateExecutionStatus


class StrictModel(BaseModel):
    """Base for versioned boundary models that reject unknown fields."""

    model_config = ConfigDict(extra="forbid")


class ExternalCapabilityKind(StrEnum):
    API = "api"
    COM = "com"
    XLL = "xll"
    FILE = "file"
    NETWORK = "network"
    DATABASE = "database"
    ADD_IN = "add_in"
    PROJECT_REFERENCE = "project_reference"


class ExternalCapability(StrictModel):
    """Typed declaration and grant for a workbook runtime dependency."""

    capability: str
    kind: ExternalCapabilityKind
    resource: str
    declared: bool = True
    granted: bool = False
    constraints: dict[str, str] = Field(default_factory=dict)

    @model_validator(mode="after")
    def grant_requires_declaration(self) -> ExternalCapability:
        if self.granted and not self.declared:
            raise ValueError("a capability cannot be granted without being declared")
        return self


class ActionKind(StrEnum):
    OPEN = "open"
    CLOSE = "close"
    SET_CELL = "set_cell"
    SET_RANGE = "set_range"
    RECALCULATE = "recalculate"
    EXECUTE_PYTHON_MACRO = "execute_python_macro"
    DISPATCH_CONTROL_EVENT = "dispatch_control_event"
    SEND_KEYBOARD_EVENT = "send_keyboard_event"
    EXPORT_PDF = "export_pdf"
    # Compatibility actions retained for existing scenario fixtures.
    INVOKE_MACRO = "invoke_macro"
    ACTIVATE_SHEET = "activate_sheet"
    COPY_SHEET = "copy_sheet"
    MOVE_SHEET = "move_sheet"
    RENAME_SHEET = "rename_sheet"
    SAVE = "save"
    SAVE_AS = "save_as"
    REOPEN = "reopen"
    CLICK_CONTROL = "click_control"
    REFRESH_DATA = "refresh_data"
    PRINT = "print"
    EXPORT = "export"


class ObservationKind(StrEnum):
    CELL_VALUE = "cell_value"
    CELL_FORMULA = "cell_formula"
    CELL_TYPE = "cell_type"
    CELL_ERROR = "cell_error"
    RANGE_VALUES = "range_values"
    SHEET_STATE = "sheet_state"
    CONTROLS_BINDINGS = "controls_bindings"
    FILES_CREATED = "files_created"
    MOCKED_CALLS = "mocked_calls"
    SCREENSHOTS = "screenshots"
    RUNTIME_ERRORS = "runtime_errors"
    # Compatibility observations retained for existing scenario fixtures.
    CELL = "cell"
    SHEETS = "sheets"
    NAMED_RANGES = "named_ranges"
    EMBEDDED_SCRIPTS = "embedded_scripts"
    CONTROLS_EVENTS = "controls_events"
    PACKAGE_HASH = "package_hash"
    ARTIFACT_INVENTORY = "artifact_inventory"


class ValueKind(StrEnum):
    EMPTY_CELL = "empty_cell"
    EMPTY_STRING = "empty_string"
    BOOLEAN = "boolean"
    NUMBER = "number"
    STRING = "string"
    ERROR = "error"
    DATE = "date"
    DATETIME = "datetime"
    ARRAY = "array"
    OBJECT = "object"


class EnvironmentManifest(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    locale: str = "en-US"
    timezone: str = "UTC"
    date_system: Literal["1900", "1904"] = "1900"
    calculation_mode: Literal["automatic", "manual", "automatic_except_tables"] = "automatic"
    iterative_calculation: bool = False
    max_iterations: int = Field(default=100, ge=1)
    max_change: float = Field(default=0.001, gt=0)
    external_workbooks: dict[str, str] = Field(default_factory=dict)
    files: dict[str, str] = Field(default_factory=dict)
    databases: dict[str, str] = Field(default_factory=dict)
    add_ins: list[str] = Field(default_factory=list)
    references: list[str] = Field(default_factory=list)
    declared_capabilities: list[str] = Field(default_factory=list)
    granted_capabilities: list[str] = Field(default_factory=list)
    typed_capabilities: list[ExternalCapability] = Field(default_factory=list)

    @model_validator(mode="after")
    def grants_must_be_declared(self) -> EnvironmentManifest:
        typed_declared = {
            capability.capability for capability in self.typed_capabilities if capability.declared
        }
        typed_granted = {
            capability.capability for capability in self.typed_capabilities if capability.granted
        }
        undeclared = (set(self.granted_capabilities) | typed_granted) - (
            set(self.declared_capabilities) | typed_declared
        )
        if undeclared:
            raise ValueError(f"granted capabilities were not declared: {sorted(undeclared)}")
        return self

    @property
    def all_granted_capabilities(self) -> set[str]:
        return set(self.granted_capabilities) | {
            capability.capability for capability in self.typed_capabilities if capability.granted
        }


class ComparisonRules(StrictModel):
    absolute_tolerance: float = Field(default=0.0, ge=0.0)
    relative_tolerance: float = Field(default=0.0, ge=0.0)
    empty_string_equals_empty_cell: bool = False
    string_case_sensitive: bool = True


class ObservationValue(StrictModel):
    """Typed observation value that never collapses spreadsheet scalar semantics."""

    kind: ValueKind
    value: Any = None
    error_type: str | None = None
    date_system: Literal["1900", "1904"] | None = None
    timezone: str | None = None
    formula: str | None = None
    cell_type: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def typed_metadata_is_complete(self) -> ObservationValue:
        if self.kind is ValueKind.ERROR and not self.error_type:
            raise ValueError("error observations require error_type")
        if self.kind in {ValueKind.DATE, ValueKind.DATETIME} and not self.date_system:
            raise ValueError("date observations require date_system")
        return self


class Action(StrictModel):
    kind: ActionKind
    parameters: dict[str, Any] = Field(default_factory=dict)
    required: bool = True


class ObservationRequest(StrictModel):
    id: str = Field(min_length=1)
    kind: ObservationKind
    selector: dict[str, Any] = Field(default_factory=dict)
    required: bool = True
    comparison: ComparisonRules = Field(default_factory=ComparisonRules)
    expected: ObservationValue | None = None


class ScenarioStep(StrictModel):
    id: str = Field(min_length=1)
    action: Action
    observations_before: list[ObservationRequest] = Field(default_factory=list)
    observations_after: list[ObservationRequest] = Field(default_factory=list)


class Scenario(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    id: str = Field(min_length=1)
    description: str = ""
    steps: list[ScenarioStep]

    @model_validator(mode="after")
    def unique_ids(self) -> Scenario:
        step_ids = [step.id for step in self.steps]
        if len(step_ids) != len(set(step_ids)):
            raise ValueError("scenario step IDs must be unique")
        observation_ids = [
            request.id
            for step in self.steps
            for request in (*step.observations_before, *step.observations_after)
        ]
        if len(observation_ids) != len(set(observation_ids)):
            raise ValueError("scenario observation IDs must be unique")
        return self


class MigrationMetadata(StrictModel):
    """Authored and independently reviewed migration acceptance metadata."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    id: str = Field(min_length=1)
    title: str = Field(min_length=1)
    source_workbook: str | None = None
    target_workbook: str | None = None
    authored_by: str = Field(min_length=1)
    reviewed_by: str = Field(min_length=1)
    requirements: list[str] = Field(min_length=1)
    oracle_policy: Literal["authored_acceptance_requirements"] = "authored_acceptance_requirements"

    @model_validator(mode="after")
    def reviewer_is_independent(self) -> MigrationMetadata:
        if self.authored_by.strip().casefold() == self.reviewed_by.strip().casefold():
            raise ValueError("acceptance author and reviewer must be independent")
        return self


class AcceptanceDefinition(StrictModel):
    """Complete public acceptance contract in YAML or JSON form."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    migration: MigrationMetadata
    environment: EnvironmentManifest = Field(default_factory=EnvironmentManifest)
    scenario: Scenario


class StepResult(StrictModel):
    step_id: str
    action: ActionKind
    status: GateExecutionStatus
    started_at: datetime
    ended_at: datetime
    observations_before: dict[str, ObservationValue] = Field(default_factory=dict)
    observations_after: dict[str, ObservationValue] = Field(default_factory=dict)
    evidence: list[str] = Field(default_factory=list)
    error: dict[str, Any] | None = None


class RuntimeIdentity(StrictModel):
    runtime_kind: str
    runtime_version: str
    executable_path: str | None = None
    executable_sha256: str | None = None
    image_reference: str | None = None
    image_digest: str | None = None
    base_image_digest: str | None = None
    architecture: str | None = None
    python_version: str | None = None
    pyuno_identity: dict[str, Any] = Field(default_factory=dict)
    package_manifest: list[dict[str, Any]] = Field(default_factory=list)
    container_configuration: dict[str, Any] = Field(default_factory=dict)
    metadata: dict[str, Any] = Field(default_factory=dict)


class RuntimeTrace(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    trace_id: str
    scenario_id: str
    runtime_role: Literal["source", "target", "fake_source", "fake_target"]
    runtime_identity: RuntimeIdentity
    environment: EnvironmentManifest
    status: GateExecutionStatus
    started_at: datetime
    ended_at: datetime
    workbook_hash_before: str
    workbook_hash_after: str | None = None
    steps: list[StepResult] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    logs: list[str] = Field(default_factory=list)
    error: dict[str, Any] | None = None


class AssertionResult(StrictModel):
    """One evaluated observation assertion."""

    step_id: str
    observation_id: str
    phase: Literal["before", "after"]
    required: bool
    status: GateExecutionStatus
    expected: ObservationValue | None = None
    actual: ObservationValue | None = None
    reason: str | None = None


class AcceptanceEvaluation(StrictModel):
    """Fail-closed evaluation of one execution trace."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    migration_id: str
    scenario_id: str
    trace_id: str
    status: GateExecutionStatus
    action_statuses: dict[str, GateExecutionStatus] = Field(default_factory=dict)
    assertions: list[AssertionResult] = Field(default_factory=list)
    required_failures: list[str] = Field(default_factory=list)


class ObservationDifference(StrictModel):
    step_id: str
    observation_id: str
    source: ObservationValue | None
    target: ObservationValue | None
    matched: bool
    reason: str | None = None


class TraceDiff(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    source_trace_id: str
    target_trace_id: str
    status: GateExecutionStatus
    differences: list[ObservationDifference] = Field(default_factory=list)
    missing_source_steps: list[str] = Field(default_factory=list)
    missing_target_steps: list[str] = Field(default_factory=list)

    @property
    def equivalent(self) -> bool:
        return self.status is GateExecutionStatus.PASSED and not any(
            not difference.matched for difference in self.differences
        )


class EvidenceBundleManifest(StrictModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    bundle_id: str
    created_at: datetime
    source_workbook_hash: str
    output_hash: str | None
    environment_manifest: str
    scenario_definition: str
    source_trace: str | None
    source_inventory: str | None = None
    target_inventories: dict[str, str] = Field(default_factory=dict)
    inventory_diffs: list[str] = Field(default_factory=list)
    target_traces: dict[str, str] = Field(default_factory=dict)
    trace_diffs: list[str] = Field(default_factory=list)
    runtime_identities: dict[str, RuntimeIdentity] = Field(default_factory=dict)
    logs: list[str] = Field(default_factory=list)
    attachments: list[str] = Field(default_factory=list)
    schema_versions: dict[str, str] = Field(default_factory=dict)
    granted_capabilities: list[str] = Field(default_factory=list)


class AcceptanceEvidenceManifest(StrictModel):
    """Content-addressed machine-readable manifest for one acceptance run."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    evidence_id: str
    created_at: datetime
    migration_id: str
    scenario_id: str
    status: GateExecutionStatus
    workbook: str
    workbook_sha256: str
    acceptance_definition: str
    execution_trace: str
    evaluation: str
    markdown_report: str
    file_sha256: dict[str, str] = Field(default_factory=dict)


class MutationOutcome(StrEnum):
    KILLED = "killed"
    SURVIVED = "survived"
    INCONCLUSIVE = "inconclusive"


class MutationCaseResult(StrictModel):
    """One isolated mutation and its public-acceptance outcome."""

    id: str
    kind: Literal["python", "formula"]
    target: str
    mutant_workbook: str
    mutant_sha256: str
    trace: str | None = None
    evaluation: str | None = None
    outcome: MutationOutcome
    reason: str


class MutationCampaign(StrictModel):
    """Aggregate mutation-test evidence."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    migration_id: str
    source_workbook_sha256: str
    status: GateExecutionStatus
    cases: list[MutationCaseResult] = Field(default_factory=list)
