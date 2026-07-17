"""Deterministic VBA semantics and procedure-by-procedure execution planning."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any, Generic, TypedDict, TypeVar

from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.runtime.object_model import Application
from xlsliberator.scenarios.models import EnvironmentManifest
from xlsliberator.vba_ir import (
    VBAParameterPassing,
    VBAProcedureIR,
    VBAProjectIR,
    VBAStatementKind,
)


class VBAValueKind(StrEnum):
    EMPTY = "empty"
    NULL = "null"
    ERROR = "error"
    BOOLEAN = "boolean"
    INTEGER = "integer"
    DOUBLE = "double"
    STRING = "string"
    DATE = "date"
    OBJECT = "object"
    ARRAY = "array"


class VBAExecutionError(RuntimeError):
    """Typed deterministic VBA runtime failure."""

    def __init__(self, number: int, description: str) -> None:
        super().__init__(description)
        self.number = number
        self.description = description


@dataclass(frozen=True)
class VBAVariant:
    kind: VBAValueKind
    value: Any = None
    error_number: int | None = None

    @classmethod
    def empty(cls) -> VBAVariant:
        return cls(VBAValueKind.EMPTY)

    @classmethod
    def null(cls) -> VBAVariant:
        return cls(VBAValueKind.NULL)

    @classmethod
    def error(cls, number: int) -> VBAVariant:
        return cls(VBAValueKind.ERROR, error_number=number)

    @classmethod
    def from_python(cls, value: Any) -> VBAVariant:
        if isinstance(value, VBAVariant):
            return value
        if value is None:
            return cls.empty()
        if isinstance(value, bool):
            return cls(VBAValueKind.BOOLEAN, value)
        if isinstance(value, int):
            return cls(VBAValueKind.INTEGER, value)
        if isinstance(value, float):
            return cls(VBAValueKind.DOUBLE, value)
        if isinstance(value, str):
            return cls(VBAValueKind.STRING, value)
        if isinstance(value, VBAArray):
            return cls(VBAValueKind.ARRAY, value)
        return cls(VBAValueKind.OBJECT, value)


def coerce_number(value: VBAVariant) -> float:
    """Implement the deterministic subset of VBA numeric coercion."""
    if value.kind is VBAValueKind.EMPTY:
        return 0.0
    if value.kind is VBAValueKind.BOOLEAN:
        return -1.0 if value.value else 0.0
    if value.kind in {VBAValueKind.INTEGER, VBAValueKind.DOUBLE}:
        return float(value.value)
    if value.kind is VBAValueKind.STRING:
        try:
            return float(str(value.value).strip())
        except ValueError as exc:
            raise VBAExecutionError(13, "Type mismatch") from exc
    if value.kind is VBAValueKind.NULL:
        raise VBAExecutionError(94, "Invalid use of Null")
    if value.kind is VBAValueKind.ERROR:
        raise VBAExecutionError(value.error_number or 13, "Error value cannot be coerced")
    raise VBAExecutionError(13, "Type mismatch")


def coerce_string(value: VBAVariant) -> str:
    if value.kind is VBAValueKind.EMPTY:
        return ""
    if value.kind is VBAValueKind.BOOLEAN:
        return "True" if value.value else "False"
    if value.kind in {VBAValueKind.INTEGER, VBAValueKind.DOUBLE, VBAValueKind.STRING}:
        return str(value.value)
    if value.kind is VBAValueKind.NULL:
        raise VBAExecutionError(94, "Invalid use of Null")
    if value.kind is VBAValueKind.ERROR:
        raise VBAExecutionError(value.error_number or 13, "Error value cannot be coerced")
    raise VBAExecutionError(13, "Type mismatch")


T = TypeVar("T")


@dataclass
class ByRefCell(Generic[T]):
    """Mutable storage location preserving VBA ByRef aliasing."""

    value: T


@dataclass
class VBAArray:
    """Bounded VBA array supporting deterministic ReDim Preserve semantics."""

    lower_bound: int
    upper_bound: int
    values: list[VBAVariant] = field(default_factory=list)

    def __post_init__(self) -> None:
        if self.upper_bound < self.lower_bound:
            raise VBAExecutionError(9, "Subscript out of range")
        expected = self.upper_bound - self.lower_bound + 1
        if not self.values:
            self.values = [VBAVariant.empty() for _ in range(expected)]
        elif len(self.values) != expected:
            raise VBAExecutionError(9, "Array bounds do not match values")

    def get(self, index: int) -> VBAVariant:
        return self.values[self._offset(index)]

    def set(self, index: int, value: VBAVariant | Any) -> None:
        self.values[self._offset(index)] = VBAVariant.from_python(value)

    def redim(self, lower_bound: int, upper_bound: int, *, preserve: bool = False) -> None:
        replacement = VBAArray(lower_bound, upper_bound)
        if preserve:
            overlap_start = max(self.lower_bound, lower_bound)
            overlap_end = min(self.upper_bound, upper_bound)
            for index in range(overlap_start, overlap_end + 1):
                replacement.set(index, self.get(index))
        self.lower_bound = replacement.lower_bound
        self.upper_bound = replacement.upper_bound
        self.values = replacement.values

    def _offset(self, index: int) -> int:
        if index < self.lower_bound or index > self.upper_bound:
            raise VBAExecutionError(9, "Subscript out of range")
        return index - self.lower_bound


@dataclass
class VBAErrorState:
    mode: str = "raise"
    last_error: VBAExecutionError | None = None

    def handle(self, error: VBAExecutionError) -> bool:
        self.last_error = error
        return self.mode == "resume_next"


@dataclass
class VBAEvent:
    name: str
    arguments: dict[str, ByRefCell[Any]] = field(default_factory=dict)
    cancelled: bool = False

    def cancel(self) -> None:
        self.cancelled = True
        if "Cancel" in self.arguments:
            self.arguments["Cancel"].value = True


@dataclass
class VBAClassInstance:
    class_name: str
    instance_id: str
    fields: dict[str, VBAVariant] = field(default_factory=dict)
    initialized: bool = False
    terminated: bool = False

    def initialize(self) -> None:
        if self.terminated:
            raise VBAExecutionError(91, "Object variable not set")
        self.initialized = True

    def terminate(self) -> None:
        self.terminated = True


class ProcedureStrategy(StrEnum):
    NATIVE_COMPATIBILITY = "native_compatibility"
    INTERPRET_TYPED_IR = "interpret_typed_ir"
    COMPILE_TARGET = "compile_target"
    TRANSLATE_PYTHON = "translate_python"
    UNAVAILABLE = "unavailable"


class DifferentialProof(BaseModel):
    model_config = ConfigDict(extra="forbid")

    procedure_id: str
    source_trace_id: str
    target_trace_id: str
    equivalent: bool


class ProcedureExecutionDecision(BaseModel):
    model_config = ConfigDict(extra="forbid")

    procedure_id: str
    strategy: ProcedureStrategy
    executable: bool
    reason: str
    source_node_id: str
    required_capabilities: list[str] = Field(default_factory=list)
    differential_proof: DifferentialProof | None = None


class VBAExecutionPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")

    schema_version: str = "1.0.0"
    project_id: str
    decisions: list[ProcedureExecutionDecision]
    missing_capabilities: list[str] = Field(default_factory=list)
    unsupported_constructs: list[str] = Field(default_factory=list)
    fully_executable: bool


class _DecisionCommon(TypedDict):
    procedure_id: str
    source_node_id: str
    required_capabilities: list[str]


_INTERPRETER_STATEMENTS = {
    VBAStatementKind.ASSIGNMENT,
    VBAStatementKind.ON_ERROR,
    VBAStatementKind.RESUME,
    VBAStatementKind.REDIM,
    VBAStatementKind.ERASE,
    VBAStatementKind.RAISE_EVENT,
    VBAStatementKind.EXIT,
    VBAStatementKind.DECLARATION,
}


class ProcedureExecutionResult(BaseModel):
    model_config = ConfigDict(extra="forbid", arbitrary_types_allowed=True)

    procedure_id: str
    return_value: Any = None
    variables: dict[str, Any] = Field(default_factory=dict)
    source_nodes_executed: list[str] = Field(default_factory=list)
    events: list[str] = Field(default_factory=list)
    error_number: int | None = None
    error_description: str | None = None


class TypedVBAInterpreter:
    """Small deterministic interpreter; unsupported statements fail closed."""

    def __init__(self, application: Application) -> None:
        self.application = application
        self.global_state: dict[str, Any] = {}

    def execute(
        self,
        procedure: VBAProcedureIR,
        arguments: dict[str, Any] | None = None,
    ) -> ProcedureExecutionResult:
        unsupported = [
            statement.kind.value
            for statement in procedure.statements
            if statement.kind not in _INTERPRETER_STATEMENTS
        ]
        if procedure.unsupported_constructs or unsupported:
            details = sorted(set(procedure.unsupported_constructs + unsupported))
            raise VBAExecutionError(
                445, "unsupported interpreter constructs: " + ", ".join(details)
            )
        frame = self._bind_parameters(procedure, arguments or {})
        error_state = VBAErrorState()
        executed: list[str] = []
        events: list[str] = []
        for statement in procedure.statements:
            executed.append(statement.source_span.node_id)
            try:
                should_exit = self._execute_statement(
                    procedure, statement.text, statement.kind, frame, error_state, events
                )
            except VBAExecutionError as exc:
                if error_state.handle(exc):
                    continue
                return ProcedureExecutionResult(
                    procedure_id=procedure.procedure_id,
                    variables=_unwrap_frame(frame),
                    source_nodes_executed=executed,
                    events=events,
                    error_number=exc.number,
                    error_description=exc.description,
                )
            if should_exit:
                break
        return ProcedureExecutionResult(
            procedure_id=procedure.procedure_id,
            return_value=_unwrap(frame.get(procedure.name)),
            variables=_unwrap_frame(frame),
            source_nodes_executed=executed,
            events=events,
            error_number=error_state.last_error.number if error_state.last_error else None,
            error_description=(
                error_state.last_error.description if error_state.last_error else None
            ),
        )

    def _bind_parameters(
        self, procedure: VBAProcedureIR, arguments: dict[str, Any]
    ) -> dict[str, Any]:
        frame: dict[str, Any] = dict(self.global_state)
        lowered_arguments = {name.lower(): value for name, value in arguments.items()}
        for parameter in procedure.parameters:
            key = parameter.name.lower()
            if key in lowered_arguments:
                value = lowered_arguments[key]
            elif parameter.optional:
                value = _evaluate_literal(parameter.default_value or "Empty")
            else:
                raise VBAExecutionError(449, f"argument not optional: {parameter.name}")
            if parameter.param_array:
                value = (
                    value
                    if isinstance(value, VBAArray)
                    else VBAArray(0, 0, [VBAVariant.from_python(value)])
                )
            if parameter.passing is VBAParameterPassing.BYREF:
                frame[parameter.name] = value if isinstance(value, ByRefCell) else ByRefCell(value)
            else:
                frame[parameter.name] = _unwrap(value)
        return frame

    def _execute_statement(
        self,
        procedure: VBAProcedureIR,
        text: str,
        kind: VBAStatementKind,
        frame: dict[str, Any],
        error_state: VBAErrorState,
        events: list[str],
    ) -> bool:
        if kind is VBAStatementKind.ON_ERROR:
            lowered = text.lower()
            if "resume next" in lowered:
                error_state.mode = "resume_next"
            elif "goto 0" in lowered:
                error_state.mode = "raise"
                error_state.last_error = None
            else:
                raise VBAExecutionError(445, f"unsupported error handler: {text}")
            return False
        if kind is VBAStatementKind.DECLARATION:
            match = _require_match(r"^(?:Dim|Static)\s+(\w+)", text)
            frame.setdefault(match.group(1), VBAVariant.empty())
            return False
        if kind is VBAStatementKind.REDIM:
            match = _require_match(
                r"^ReDim\s+(Preserve\s+)?(\w+)\s*\(\s*(?:(-?\d+)\s+To\s+)?(-?\d+)\s*\)$",
                text,
            )
            preserve = bool(match.group(1))
            name = match.group(2)
            lower = int(match.group(3) or 0)
            upper = int(match.group(4))
            existing = _unwrap(frame.get(name))
            if preserve and isinstance(existing, VBAArray):
                existing.redim(lower, upper, preserve=True)
                _assign(frame, name, existing)
            else:
                _assign(frame, name, VBAArray(lower, upper))
            return False
        if kind is VBAStatementKind.ERASE:
            match = _require_match(r"^Erase\s+(\w+)$", text)
            _assign(frame, match.group(1), VBAVariant.empty())
            return False
        if kind is VBAStatementKind.RAISE_EVENT:
            match = _require_match(r"^RaiseEvent\s+(\w+)", text)
            event_name = match.group(1)
            events.append(event_name)
            self.application.active_workbook.emit_event(event_name)
            return False
        if kind is VBAStatementKind.EXIT:
            return True
        if kind is VBAStatementKind.RESUME:
            if error_state.last_error is None:
                raise VBAExecutionError(20, "Resume without error")
            error_state.last_error = None
            return False
        if kind is VBAStatementKind.ASSIGNMENT:
            self._assignment(procedure, text, frame)
            return False
        raise VBAExecutionError(445, f"unsupported interpreter statement: {kind.value}")

    def _assignment(self, procedure: VBAProcedureIR, text: str, frame: dict[str, Any]) -> None:
        match = _require_match(r"^(?:Let\s+|Set\s+)?(.+?)\s*=\s*(.+)$", text)
        target, expression = match.groups()
        value = _evaluate_expression(expression.strip(), frame)
        if range_match := re.match(
            r'^(?:ActiveSheet\.)?Range\s*\(\s*"([^"]+)"\s*\)\.(?:Value|Value2)$',
            target.strip(),
            re.IGNORECASE,
        ):
            self.application.active_sheet.range(range_match.group(1)).value = _unwrap(value)
            return
        if cells_match := re.match(
            r"^(?:ActiveSheet\.)?Cells\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)\.(?:Value|Value2)$",
            target.strip(),
            re.IGNORECASE,
        ):
            self.application.active_sheet.cells(
                int(cells_match.group(1)), int(cells_match.group(2))
            ).value = _unwrap(value)
            return
        if re.fullmatch(r"[A-Za-z_]\w*", target.strip()):
            _assign(frame, target.strip(), value)
            if target.strip().lower() not in {
                parameter.name.lower() for parameter in procedure.parameters
            }:
                self.global_state[target.strip()] = frame[target.strip()]
            return
        raise VBAExecutionError(445, f"unsupported assignment target: {target.strip()}")


def build_execution_plan(
    project: VBAProjectIR,
    environment: EnvironmentManifest,
    *,
    differential_proofs: list[DifferentialProof] | None = None,
    preferred_strategy: ProcedureStrategy | None = None,
) -> VBAExecutionPlan:
    """Select an evidence-backed strategy independently for every procedure."""
    granted = environment.all_granted_capabilities
    missing = sorted(set(project.required_capabilities) - granted)
    proofs = {proof.procedure_id: proof for proof in differential_proofs or []}
    decisions: list[ProcedureExecutionDecision] = []
    for module in project.modules:
        module_capabilities = sorted(
            {dependency.capability for dependency in module.external_dependencies}
        )
        for procedure in module.procedures:
            procedure_capabilities = {
                dependency.capability for dependency in procedure.external_dependencies
            }
            procedure_missing = sorted(
                (set(module_capabilities) | procedure_capabilities) - granted
            )
            decision = _decide_procedure(
                procedure,
                procedure_missing,
                proofs.get(procedure.procedure_id),
                granted,
                preferred_strategy,
            )
            decisions.append(decision)
    return VBAExecutionPlan(
        project_id=project.project_id,
        decisions=decisions,
        missing_capabilities=missing,
        unsupported_constructs=project.unsupported_constructs,
        fully_executable=bool(decisions) and all(decision.executable for decision in decisions),
    )


def _decide_procedure(
    procedure: VBAProcedureIR,
    missing: list[str],
    proof: DifferentialProof | None,
    granted: set[str],
    preferred: ProcedureStrategy | None,
) -> ProcedureExecutionDecision:
    common: _DecisionCommon = {
        "procedure_id": procedure.procedure_id,
        "source_node_id": procedure.source_span.node_id,
        "required_capabilities": missing,
    }
    if missing:
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.UNAVAILABLE,
            executable=False,
            reason=f"missing capabilities: {', '.join(missing)}",
        )
    if procedure.unsupported_constructs:
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.UNAVAILABLE,
            executable=False,
            reason="unsupported constructs: " + ", ".join(procedure.unsupported_constructs),
        )
    if preferred is ProcedureStrategy.TRANSLATE_PYTHON:
        if proof is None or not proof.equivalent:
            return ProcedureExecutionDecision(
                **common,
                strategy=ProcedureStrategy.UNAVAILABLE,
                executable=False,
                reason="Python translation lacks equivalent source/target differential proof",
                differential_proof=proof,
            )
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.TRANSLATE_PYTHON,
            executable=True,
            reason="Python candidate has equivalent source/target differential proof",
            differential_proof=proof,
        )
    if "vba.native_compatibility" in granted:
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.NATIVE_COMPATIBILITY,
            executable=True,
            reason="native compatibility capability is explicitly granted",
        )
    if all(statement.kind in _INTERPRETER_STATEMENTS for statement in procedure.statements):
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.INTERPRET_TYPED_IR,
            executable=True,
            reason="all statements are in the deterministic interpreter subset",
        )
    if "vba.target_compiler" in granted:
        return ProcedureExecutionDecision(
            **common,
            strategy=ProcedureStrategy.COMPILE_TARGET,
            executable=True,
            reason="typed target compiler capability is explicitly granted",
        )
    return ProcedureExecutionDecision(
        **common,
        strategy=ProcedureStrategy.UNAVAILABLE,
        executable=False,
        reason="no implemented deterministic strategy covers all statements",
    )


def _evaluate_literal(text: str) -> Any:
    value = text.strip()
    lowered = value.lower()
    if lowered == "empty":
        return VBAVariant.empty()
    if lowered == "null":
        return VBAVariant.null()
    if lowered == "true":
        return True
    if lowered == "false":
        return False
    if re.fullmatch(r"[-+]?\d+", value):
        return int(value)
    if re.fullmatch(r"[-+]?(?:\d+\.\d*|\d*\.\d+)", value):
        return float(value)
    if len(value) >= 2 and value.startswith('"') and value.endswith('"'):
        return value[1:-1].replace('""', '"')
    raise VBAExecutionError(445, f"unsupported literal: {text}")


def _evaluate_expression(text: str, frame: dict[str, Any]) -> Any:
    value = text.strip()
    if re.fullmatch(r"[A-Za-z_]\w*", value) and value in frame:
        return _unwrap(frame[value])
    for operator in ("+", "-"):
        match = re.fullmatch(r"(.+?)\s*" + re.escape(operator) + r"\s*(.+)", value)
        if match:
            left = VBAVariant.from_python(_evaluate_expression(match.group(1), frame))
            right = VBAVariant.from_python(_evaluate_expression(match.group(2), frame))
            result = coerce_number(left) + coerce_number(right)
            if operator == "-":
                result = coerce_number(left) - coerce_number(right)
            return result
    return _evaluate_literal(value)


def _assign(frame: dict[str, Any], name: str, value: Any) -> None:
    existing = frame.get(name)
    if isinstance(existing, ByRefCell):
        existing.value = _unwrap(value)
    else:
        frame[name] = value


def _unwrap(value: Any) -> Any:
    if isinstance(value, ByRefCell):
        return _unwrap(value.value)
    if isinstance(value, VBAVariant):
        return value.value if value.kind not in {VBAValueKind.EMPTY, VBAValueKind.NULL} else value
    return value


def _unwrap_frame(frame: dict[str, Any]) -> dict[str, Any]:
    return {name: _unwrap(value) for name, value in frame.items()}


def _require_match(pattern: str, text: str) -> re.Match[str]:
    match = re.match(pattern, text, re.IGNORECASE)
    if match is None:
        raise VBAExecutionError(445, f"unsupported syntax: {text}")
    return match
