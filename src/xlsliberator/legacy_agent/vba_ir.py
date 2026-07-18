"""Deprecated semantic IR used only by the embedded legacy translator."""

from __future__ import annotations

from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

VBA_PROJECT_IR_SCHEMA_VERSION = "1.0.0"


class StrictIRModel(BaseModel):
    """Reject undeclared boundary fields in persisted VBA IR."""

    model_config = ConfigDict(extra="forbid")


class SourceSpan(StrictIRModel):
    """Stable source location attached to every executable/declarative node."""

    node_id: str
    module_name: str
    start_offset: int = Field(ge=0)
    end_offset: int = Field(ge=0)
    start_line: int = Field(ge=1)
    start_column: int = Field(ge=1)
    end_line: int = Field(ge=1)
    end_column: int = Field(ge=1)
    source_sha256: str


class VBAModuleKind(StrEnum):
    STANDARD = "standard"
    CLASS = "class"
    DOCUMENT = "document"
    USERFORM = "userform"


class VBAVisibility(StrEnum):
    PUBLIC = "public"
    PRIVATE = "private"
    FRIEND = "friend"


class VBAProcedureKind(StrEnum):
    SUB = "sub"
    FUNCTION = "function"
    PROPERTY_GET = "property_get"
    PROPERTY_LET = "property_let"
    PROPERTY_SET = "property_set"


class VBAParameterPassing(StrEnum):
    BYREF = "byref"
    BYVAL = "byval"


class VBAStatementKind(StrEnum):
    ASSIGNMENT = "assignment"
    CALL = "call"
    IF = "if"
    SELECT = "select"
    FOR = "for"
    FOR_EACH = "for_each"
    DO = "do"
    WHILE = "while"
    WITH = "with"
    ON_ERROR = "on_error"
    RESUME = "resume"
    GOTO = "goto"
    REDIM = "redim"
    ERASE = "erase"
    RAISE_EVENT = "raise_event"
    EXIT = "exit"
    DECLARATION = "declaration"
    RAW = "raw"


class VBAExpressionKind(StrEnum):
    LITERAL = "literal"
    IDENTIFIER = "identifier"
    MEMBER_ACCESS = "member_access"
    DEFAULT_MEMBER = "default_member"
    CALL = "call"
    NEW_OBJECT = "new_object"
    ARRAY_ACCESS = "array_access"
    BINARY = "binary"
    UNARY = "unary"
    RAW = "raw"


class VBAExternalDependencyKind(StrEnum):
    API = "api"
    COM = "com"
    XLL = "xll"
    FILE = "file"
    NETWORK = "network"
    DATABASE = "database"
    ADD_IN = "add_in"
    PROJECT_REFERENCE = "project_reference"


class VBAProjectReference(StrictIRModel):
    name: str
    guid: str | None = None
    major: int | None = Field(default=None, ge=0)
    minor: int | None = Field(default=None, ge=0)
    path: str | None = None
    built_in: bool = False
    capability: str


class VBAConditionalConstant(StrictIRModel):
    name: str
    value: str
    source_span: SourceSpan


class VBAConditionalBlock(StrictIRModel):
    condition: str
    branches: list[str] = Field(default_factory=list)
    source_span: SourceSpan


class VBAParameterIR(StrictIRModel):
    name: str
    type_name: str = "Variant"
    passing: VBAParameterPassing = VBAParameterPassing.BYREF
    optional: bool = False
    param_array: bool = False
    default_value: str | None = None
    is_array: bool = False
    source_span: SourceSpan

    @model_validator(mode="after")
    def param_array_contract(self) -> VBAParameterIR:
        if self.param_array and not self.is_array:
            raise ValueError("ParamArray parameters must be arrays")
        return self


class VBAVariableDeclaration(StrictIRModel):
    name: str
    type_name: str = "Variant"
    visibility: VBAVisibility = VBAVisibility.PRIVATE
    is_array: bool = False
    bounds: list[str] = Field(default_factory=list)
    with_events: bool = False
    is_static: bool = False
    source_span: SourceSpan


class VBAConstantDeclaration(StrictIRModel):
    name: str
    type_name: str = "Variant"
    value_expression: str
    visibility: VBAVisibility = VBAVisibility.PRIVATE
    source_span: SourceSpan


class VBAEnumMember(StrictIRModel):
    name: str
    value_expression: str | None = None
    source_span: SourceSpan


class VBAEnumDeclaration(StrictIRModel):
    name: str
    visibility: VBAVisibility
    members: list[VBAEnumMember]
    source_span: SourceSpan


class VBAUDTField(StrictIRModel):
    name: str
    type_name: str = "Variant"
    is_array: bool = False
    source_span: SourceSpan


class VBAUserDefinedType(StrictIRModel):
    name: str
    visibility: VBAVisibility
    fields: list[VBAUDTField]
    source_span: SourceSpan


class VBAExpressionIR(StrictIRModel):
    kind: VBAExpressionKind
    text: str
    inferred_type: str | None = None
    object_reference: str | None = None
    member_name: str | None = None
    uses_default_member: bool = False
    source_span: SourceSpan


class VBAStatementIR(StrictIRModel):
    kind: VBAStatementKind
    text: str
    expressions: list[VBAExpressionIR] = Field(default_factory=list)
    labels: list[str] = Field(default_factory=list)
    error_mode: str | None = None
    source_span: SourceSpan


class VBAExternalDependency(StrictIRModel):
    dependency_id: str
    kind: VBAExternalDependencyKind
    name: str
    capability: str
    required: bool = True
    details: dict[str, str] = Field(default_factory=dict)
    source_span: SourceSpan


class VBAProcedureIR(StrictIRModel):
    procedure_id: str
    name: str
    kind: VBAProcedureKind
    visibility: VBAVisibility
    parameters: list[VBAParameterIR] = Field(default_factory=list)
    return_type: str | None = None
    statements: list[VBAStatementIR] = Field(default_factory=list)
    local_variables: list[VBAVariableDeclaration] = Field(default_factory=list)
    is_static: bool = False
    is_event_handler: bool = False
    cancels_event: bool = False
    mutates_global_state: bool = False
    calls: list[str] = Field(default_factory=list)
    external_dependencies: list[VBAExternalDependency] = Field(default_factory=list)
    unsupported_constructs: list[str] = Field(default_factory=list)
    source_span: SourceSpan


class VBAModuleIR(StrictIRModel):
    module_id: str
    name: str
    kind: VBAModuleKind
    source_sha256: str
    source_code: str
    attributes: dict[str, str] = Field(default_factory=dict)
    conditional_constants: list[VBAConditionalConstant] = Field(default_factory=list)
    conditional_blocks: list[VBAConditionalBlock] = Field(default_factory=list)
    variables: list[VBAVariableDeclaration] = Field(default_factory=list)
    constants: list[VBAConstantDeclaration] = Field(default_factory=list)
    enums: list[VBAEnumDeclaration] = Field(default_factory=list)
    user_defined_types: list[VBAUserDefinedType] = Field(default_factory=list)
    procedures: list[VBAProcedureIR] = Field(default_factory=list)
    external_dependencies: list[VBAExternalDependency] = Field(default_factory=list)
    has_global_state: bool = False
    lifetime: Literal["project", "instance", "document", "form"]
    source_span: SourceSpan


class VBAProjectIR(StrictIRModel):
    schema_version: Literal["1.0.0"] = "1.0.0"
    parser_name: str = "xlsliberator-deterministic-vba-parser"
    parser_version: str = "1.0.0"
    project_id: str
    name: str
    references: list[VBAProjectReference] = Field(default_factory=list)
    conditional_compilation_arguments: dict[str, str] = Field(default_factory=dict)
    modules: list[VBAModuleIR]
    external_dependencies: list[VBAExternalDependency] = Field(default_factory=list)
    required_capabilities: list[str] = Field(default_factory=list)
    unsupported_constructs: list[str] = Field(default_factory=list)
    source_map: dict[str, SourceSpan] = Field(default_factory=dict)

    @model_validator(mode="after")
    def stable_identity_contract(self) -> VBAProjectIR:
        node_ids = list(self.source_map)
        if len(node_ids) != len(set(node_ids)):
            raise ValueError("source-map node IDs must be unique")
        if any(node_id != span.node_id for node_id, span in self.source_map.items()):
            raise ValueError("source-map keys must match span node IDs")
        return self
