"""Deprecated parser used only by the embedded legacy translator.

The parser intentionally preserves raw source text alongside typed nodes. It
never treats a successfully tokenized but unsupported construct as executable.
"""

from __future__ import annotations

import hashlib
import re
from collections.abc import Mapping, Sequence
from typing import Any, Literal, cast

from pydantic import BaseModel

from xlsliberator.legacy_agent.vba_ir import (
    SourceSpan,
    VBAConditionalBlock,
    VBAConditionalConstant,
    VBAConstantDeclaration,
    VBAEnumDeclaration,
    VBAEnumMember,
    VBAExpressionIR,
    VBAExpressionKind,
    VBAExternalDependency,
    VBAExternalDependencyKind,
    VBAModuleIR,
    VBAModuleKind,
    VBAParameterIR,
    VBAParameterPassing,
    VBAProcedureIR,
    VBAProcedureKind,
    VBAProjectIR,
    VBAProjectReference,
    VBAStatementIR,
    VBAStatementKind,
    VBAUDTField,
    VBAUserDefinedType,
    VBAVariableDeclaration,
    VBAVisibility,
)

PARSER_VERSION = "1.0.0"

_PROCEDURE_START = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?(?:(Static)\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)\s*"
    r"\((.*?)\)\s*(?:As\s+([A-Za-z_][\w.]*))?\s*$",
    re.IGNORECASE,
)
_PROCEDURE_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\s*$", re.IGNORECASE)
_ATTRIBUTE = re.compile(r"^\s*Attribute\s+([\w.]+)\s*=\s*(.*?)\s*$", re.IGNORECASE)
_VARIABLE = re.compile(
    r"^\s*(?:(Public|Private|Friend|Dim|Static)\s+)(WithEvents\s+)?"
    r"([A-Za-z_]\w*)\s*(\([^)]*\))?\s*(?:As\s+([A-Za-z_][\w.]*))?",
    re.IGNORECASE,
)
_CONSTANT = re.compile(
    r"^\s*(?:(Public|Private|Friend)\s+)?Const\s+([A-Za-z_]\w*)\s*"
    r"(?:As\s+([A-Za-z_][\w.]*))?\s*=\s*(.+)$",
    re.IGNORECASE,
)
_DECLARE = re.compile(
    r"^\s*(?:(Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?"
    r"(?:Sub|Function)\s+([A-Za-z_]\w*)\s+Lib\s+\"([^\"]+)\"",
    re.IGNORECASE,
)


class VBAParseError(ValueError):
    """Source cannot be represented without inventing structure."""


def parse_vba_project(
    name: str,
    modules: Mapping[str, str],
    *,
    module_kinds: Mapping[str, VBAModuleKind | str] | None = None,
    references: Sequence[VBAProjectReference | Mapping[str, Any]] = (),
    conditional_compilation_arguments: Mapping[str, str] | None = None,
) -> VBAProjectIR:
    """Parse a complete VBA project into deterministic, source-mapped IR."""
    if not modules:
        raise VBAParseError("VBA project contains no modules")
    normalized = {
        module_name: source.replace("\r\n", "\n").replace("\r", "\n")
        for module_name, source in modules.items()
    }
    project_seed = "\0".join(
        f"{module_name}\0{_sha256(source)}" for module_name, source in sorted(normalized.items())
    )
    project_id = f"vba-project:{_sha256(name + chr(0) + project_seed)[:24]}"
    parsed_modules = [
        _parse_module(
            project_id,
            module_name,
            source,
            _coerce_module_kind(module_name, source, (module_kinds or {}).get(module_name)),
        )
        for module_name, source in sorted(normalized.items())
    ]
    parsed_references = [
        item if isinstance(item, VBAProjectReference) else VBAProjectReference.model_validate(item)
        for item in references
    ]
    dependencies = _deduplicate_dependencies(
        dependency for module in parsed_modules for dependency in module.external_dependencies
    )
    reference_capabilities = {reference.capability for reference in parsed_references}
    required_capabilities = sorted(
        reference_capabilities | {dependency.capability for dependency in dependencies}
    )
    unsupported = sorted(
        {
            construct
            for module in parsed_modules
            for procedure in module.procedures
            for construct in procedure.unsupported_constructs
        }
    )
    project = VBAProjectIR(
        project_id=project_id,
        name=name,
        references=parsed_references,
        conditional_compilation_arguments=dict(conditional_compilation_arguments or {}),
        modules=parsed_modules,
        external_dependencies=dependencies,
        required_capabilities=required_capabilities,
        unsupported_constructs=unsupported,
    )
    project.source_map = _source_map(project)
    return project


def _parse_module(
    project_id: str,
    name: str,
    source: str,
    kind: VBAModuleKind,
) -> VBAModuleIR:
    source_sha = _sha256(source)
    offsets = _line_offsets(source)
    module_span = _span(project_id, name, "module", name, source, offsets, 0, len(source))
    attributes: dict[str, str] = {}
    constants: list[VBAConstantDeclaration] = []
    variables: list[VBAVariableDeclaration] = []
    enums: list[VBAEnumDeclaration] = []
    udts: list[VBAUserDefinedType] = []
    conditional_constants: list[VBAConditionalConstant] = []
    conditional_blocks: list[VBAConditionalBlock] = []
    procedures: list[VBAProcedureIR] = []
    dependencies: list[VBAExternalDependency] = []

    lines = source.splitlines(keepends=True)
    index = 0
    while index < len(lines):
        raw = lines[index]
        line = raw.rstrip("\r\n")
        start = offsets[index]
        end = start + len(line)
        if match := _ATTRIBUTE.match(line):
            attributes[match.group(1)] = match.group(2).strip().strip('"')
        elif match := _PROCEDURE_START.match(line):
            procedure, next_index = _parse_procedure(
                project_id, name, source, offsets, lines, index, match
            )
            procedures.append(procedure)
            dependencies.extend(procedure.external_dependencies)
            index = next_index
            continue
        elif match := re.match(r"^\s*#Const\s+(\w+)\s*=\s*(.+)$", line, re.IGNORECASE):
            conditional_constants.append(
                VBAConditionalConstant(
                    name=match.group(1),
                    value=match.group(2).strip(),
                    source_span=_span(
                        project_id,
                        name,
                        "conditional-const",
                        match.group(1),
                        source,
                        offsets,
                        start,
                        end,
                    ),
                )
            )
        elif match := re.match(r"^\s*#If\s+(.+?)\s+Then\s*$", line, re.IGNORECASE):
            block, next_index = _parse_conditional_block(
                project_id, name, source, offsets, lines, index, match.group(1)
            )
            conditional_blocks.append(block)
            index = next_index
            continue
        elif match := re.match(
            r"^\s*(?:(Public|Private|Friend)\s+)?Enum\s+(\w+)", line, re.IGNORECASE
        ):
            enum, next_index = _parse_enum(project_id, name, source, offsets, lines, index, match)
            enums.append(enum)
            index = next_index
            continue
        elif match := re.match(
            r"^\s*(?:(Public|Private|Friend)\s+)?Type\s+(\w+)", line, re.IGNORECASE
        ):
            udt, next_index = _parse_udt(project_id, name, source, offsets, lines, index, match)
            udts.append(udt)
            index = next_index
            continue
        elif match := _CONSTANT.match(line):
            constants.append(
                VBAConstantDeclaration(
                    name=match.group(2),
                    type_name=match.group(3) or "Variant",
                    value_expression=match.group(4).strip(),
                    visibility=_visibility(match.group(1)),
                    source_span=_span(
                        project_id, name, "constant", match.group(2), source, offsets, start, end
                    ),
                )
            )
        elif match := _VARIABLE.match(line):
            variables.append(_variable(project_id, name, source, offsets, start, end, match))
        dependencies.extend(
            _dependencies_for_line(project_id, name, source, offsets, start, end, line)
        )
        index += 1

    lifetime = cast(
        Literal["project", "instance", "document", "form"],
        {
            VBAModuleKind.STANDARD: "project",
            VBAModuleKind.CLASS: "instance",
            VBAModuleKind.DOCUMENT: "document",
            VBAModuleKind.USERFORM: "form",
        }[kind],
    )
    for line_index, raw_line in enumerate(lines):
        line = raw_line.rstrip("\r\n")
        line_start = offsets[line_index]
        dependencies.extend(
            _dependencies_for_line(
                project_id,
                name,
                source,
                offsets,
                line_start,
                line_start + len(line),
                line,
            )
        )
    return VBAModuleIR(
        module_id=module_span.node_id,
        name=name,
        kind=kind,
        source_sha256=source_sha,
        source_code=source,
        attributes=attributes,
        conditional_constants=conditional_constants,
        conditional_blocks=conditional_blocks,
        variables=variables,
        constants=constants,
        enums=enums,
        user_defined_types=udts,
        procedures=procedures,
        external_dependencies=_deduplicate_dependencies(dependencies),
        has_global_state=bool(variables),
        lifetime=lifetime,
        source_span=module_span,
    )


def _parse_procedure(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    lines: list[str],
    start_index: int,
    match: re.Match[str],
) -> tuple[VBAProcedureIR, int]:
    end_index = start_index + 1
    while end_index < len(lines) and not _PROCEDURE_END.match(lines[end_index].rstrip("\r\n")):
        end_index += 1
    if end_index >= len(lines):
        raise VBAParseError(f"unterminated procedure {module_name}.{match.group(4)}")
    start = offsets[start_index]
    end = offsets[end_index] + len(lines[end_index].rstrip("\r\n"))
    name = match.group(4)
    span = _span(project_id, module_name, "procedure", name, source, offsets, start, end)
    parameters = _parse_parameters(
        project_id, module_name, source, offsets, start, match.group(5), name
    )
    statements: list[VBAStatementIR] = []
    locals_: list[VBAVariableDeclaration] = []
    dependencies: list[VBAExternalDependency] = []
    unsupported: set[str] = set()
    calls: set[str] = set()
    for index in range(start_index + 1, end_index):
        raw_line = lines[index].rstrip("\r\n")
        stripped = _strip_comment(raw_line).strip()
        if not stripped:
            continue
        line_start = offsets[index]
        line_end = line_start + len(raw_line)
        statement = _statement(
            project_id, module_name, source, offsets, line_start, line_end, stripped
        )
        statements.append(statement)
        if match_variable := _VARIABLE.match(stripped):
            locals_.append(
                _variable(
                    project_id,
                    module_name,
                    source,
                    offsets,
                    line_start,
                    line_end,
                    match_variable,
                )
            )
        dependencies.extend(
            _dependencies_for_line(
                project_id, module_name, source, offsets, line_start, line_end, stripped
            )
        )
        calls.update(_called_names(stripped))
        for construct, pattern in {
            "gosub": r"\bGoSub\b",
            "implements": r"\bImplements\b",
            "address_of": r"\bAddressOf\b",
            "line_numbers": r"^\d+\s+",
        }.items():
            if re.search(pattern, stripped, re.IGNORECASE):
                unsupported.add(construct)
    kind = _procedure_kind(match.group(3))
    parameter_names = {parameter.name.lower() for parameter in parameters}
    event_handler = "_" in name and (
        module_name.lower().startswith(("sheet", "thisworkbook", "userform"))
        or name.lower().endswith(("_click", "_change", "_open", "_initialize"))
    )
    return (
        VBAProcedureIR(
            procedure_id=span.node_id,
            name=name,
            kind=kind,
            visibility=_visibility(match.group(1)),
            parameters=parameters,
            return_type=match.group(6)
            or ("Variant" if kind is VBAProcedureKind.FUNCTION else None),
            statements=statements,
            local_variables=locals_,
            is_static=bool(match.group(2)),
            is_event_handler=event_handler,
            cancels_event="cancel" in parameter_names,
            mutates_global_state=any(
                _statement_mutates_global(statement) for statement in statements
            ),
            calls=sorted(calls),
            external_dependencies=_deduplicate_dependencies(dependencies),
            unsupported_constructs=sorted(unsupported),
            source_span=span,
        ),
        end_index + 1,
    )


def _parse_parameters(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    procedure_start: int,
    text: str,
    procedure_name: str,
) -> list[VBAParameterIR]:
    parameters: list[VBAParameterIR] = []
    for position, raw in enumerate(_split_commas(text)):
        value = raw.strip()
        if not value:
            continue
        match = re.match(
            r"^(Optional\s+)?(ParamArray\s+)?(ByRef\s+|ByVal\s+)?"
            r"([A-Za-z_]\w*)\s*(\([^)]*\))?\s*(?:As\s+([A-Za-z_][\w.]*))?"
            r"\s*(?:=\s*(.+))?$",
            value,
            re.IGNORECASE,
        )
        if not match:
            raise VBAParseError(f"unsupported parameter declaration in {procedure_name}: {value}")
        relative = source.find(raw, procedure_start)
        start = relative if relative >= 0 else procedure_start
        end = start + len(raw)
        passing = (
            VBAParameterPassing.BYVAL
            if (match.group(3) or "").strip().lower() == "byval"
            else VBAParameterPassing.BYREF
        )
        param_array = bool(match.group(2))
        parameters.append(
            VBAParameterIR(
                name=match.group(4),
                type_name=match.group(6) or "Variant",
                passing=passing,
                optional=bool(match.group(1)) or param_array,
                param_array=param_array,
                default_value=match.group(7).strip() if match.group(7) else None,
                is_array=bool(match.group(5)) or param_array,
                source_span=_span(
                    project_id,
                    module_name,
                    "parameter",
                    f"{procedure_name}:{position}:{match.group(4)}",
                    source,
                    offsets,
                    start,
                    end,
                ),
            )
        )
    return parameters


def _statement(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    start: int,
    end: int,
    text: str,
) -> VBAStatementIR:
    lowered = text.lower()
    patterns = [
        (VBAStatementKind.ON_ERROR, r"^on\s+error\b"),
        (VBAStatementKind.RESUME, r"^resume\b"),
        (VBAStatementKind.GOTO, r"^goto\b"),
        (VBAStatementKind.REDIM, r"^redim\b"),
        (VBAStatementKind.ERASE, r"^erase\b"),
        (VBAStatementKind.RAISE_EVENT, r"^raiseevent\b"),
        (VBAStatementKind.IF, r"^(?:if|elseif)\b"),
        (VBAStatementKind.SELECT, r"^(?:select\s+case|case\b)"),
        (VBAStatementKind.FOR_EACH, r"^for\s+each\b"),
        (VBAStatementKind.FOR, r"^for\b"),
        (VBAStatementKind.DO, r"^(?:do|loop)\b"),
        (VBAStatementKind.WHILE, r"^(?:while|wend)\b"),
        (VBAStatementKind.WITH, r"^(?:with|end\s+with)\b"),
        (VBAStatementKind.EXIT, r"^exit\b"),
        (VBAStatementKind.DECLARATION, r"^(?:dim|static)\b"),
        (VBAStatementKind.CALL, r"^(?:call\s+)?[A-Za-z_]\w*\s*(?:\(|$)"),
    ]
    kind = VBAStatementKind.ASSIGNMENT if "=" in text else VBAStatementKind.RAW
    for candidate, pattern in patterns:
        if candidate is VBAStatementKind.CALL and "=" in text:
            continue
        if re.search(pattern, lowered, re.IGNORECASE):
            kind = candidate
            break
    span = _span(
        project_id, module_name, "statement", f"{start}:{kind.value}", source, offsets, start, end
    )
    expressions = _expressions(project_id, module_name, source, offsets, start, end, text)
    error_mode = text if kind in {VBAStatementKind.ON_ERROR, VBAStatementKind.RESUME} else None
    labels = [
        match.group(1)
        for match in re.finditer(r"\b(?:GoTo|Resume)\s+([A-Za-z_]\w*)", text, re.IGNORECASE)
    ]
    return VBAStatementIR(
        kind=kind,
        text=text,
        expressions=expressions,
        labels=labels,
        error_mode=error_mode,
        source_span=span,
    )


def _expressions(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    start: int,
    end: int,
    text: str,
) -> list[VBAExpressionIR]:
    kind = VBAExpressionKind.RAW
    member_name = None
    object_reference = None
    default_member = False
    if re.fullmatch(
        r'\s*(?:"(?:[^"]|"")*"|[-+]?\d+(?:\.\d+)?|True|False|Empty|Null)\s*', text, re.IGNORECASE
    ):
        kind = VBAExpressionKind.LITERAL
    elif re.search(r"\bNew\s+[A-Za-z_]\w*", text, re.IGNORECASE):
        kind = VBAExpressionKind.NEW_OBJECT
    elif match := re.search(r"\b([A-Za-z_]\w*)\.([A-Za-z_]\w*)", text):
        kind = VBAExpressionKind.MEMBER_ACCESS
        object_reference, member_name = match.groups()
    elif re.search(r"\b[A-Za-z_]\w*\s*\([^)]*\)", text):
        kind = VBAExpressionKind.CALL
        default_member = not bool(
            re.match(r"^\s*(?:Call|If|For|Do|While|Select|ReDim)\b", text, re.IGNORECASE)
        )
        if default_member:
            kind = VBAExpressionKind.DEFAULT_MEMBER
    span = _span(
        project_id, module_name, "expression", f"{start}:{kind.value}", source, offsets, start, end
    )
    return [
        VBAExpressionIR(
            kind=kind,
            text=text,
            object_reference=object_reference,
            member_name=member_name,
            uses_default_member=default_member,
            source_span=span,
        )
    ]


def _parse_conditional_block(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    lines: list[str],
    start_index: int,
    condition: str,
) -> tuple[VBAConditionalBlock, int]:
    branches = [condition.strip()]
    index = start_index + 1
    while index < len(lines) and not re.match(r"^\s*#End\s+If", lines[index], re.IGNORECASE):
        if match := re.match(r"^\s*#ElseIf\s+(.+?)\s+Then", lines[index], re.IGNORECASE):
            branches.append(match.group(1).strip())
        elif re.match(r"^\s*#Else\b", lines[index], re.IGNORECASE):
            branches.append("else")
        index += 1
    if index >= len(lines):
        raise VBAParseError(f"unterminated conditional compilation block in {module_name}")
    start = offsets[start_index]
    end = offsets[index] + len(lines[index].rstrip("\r\n"))
    return (
        VBAConditionalBlock(
            condition=condition.strip(),
            branches=branches,
            source_span=_span(
                project_id, module_name, "conditional", condition, source, offsets, start, end
            ),
        ),
        index + 1,
    )


def _parse_enum(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    lines: list[str],
    start_index: int,
    match: re.Match[str],
) -> tuple[VBAEnumDeclaration, int]:
    index = start_index + 1
    members: list[VBAEnumMember] = []
    while index < len(lines) and not re.match(r"^\s*End\s+Enum", lines[index], re.IGNORECASE):
        text = _strip_comment(lines[index].rstrip("\r\n")).strip()
        if member := re.match(r"^([A-Za-z_]\w*)\s*(?:=\s*(.+))?$", text):
            start = offsets[index]
            members.append(
                VBAEnumMember(
                    name=member.group(1),
                    value_expression=member.group(2).strip() if member.group(2) else None,
                    source_span=_span(
                        project_id,
                        module_name,
                        "enum-member",
                        member.group(1),
                        source,
                        offsets,
                        start,
                        start + len(text),
                    ),
                )
            )
        index += 1
    if index >= len(lines):
        raise VBAParseError(f"unterminated Enum {match.group(2)}")
    start = offsets[start_index]
    end = offsets[index] + len(lines[index].rstrip("\r\n"))
    return (
        VBAEnumDeclaration(
            name=match.group(2),
            visibility=_visibility(match.group(1)),
            members=members,
            source_span=_span(
                project_id, module_name, "enum", match.group(2), source, offsets, start, end
            ),
        ),
        index + 1,
    )


def _parse_udt(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    lines: list[str],
    start_index: int,
    match: re.Match[str],
) -> tuple[VBAUserDefinedType, int]:
    index = start_index + 1
    fields: list[VBAUDTField] = []
    while index < len(lines) and not re.match(r"^\s*End\s+Type", lines[index], re.IGNORECASE):
        text = _strip_comment(lines[index].rstrip("\r\n")).strip()
        if field := re.match(
            r"^([A-Za-z_]\w*)\s*(\([^)]*\))?\s*(?:As\s+([A-Za-z_][\w.]*))?", text, re.IGNORECASE
        ):
            start = offsets[index]
            fields.append(
                VBAUDTField(
                    name=field.group(1),
                    type_name=field.group(3) or "Variant",
                    is_array=bool(field.group(2)),
                    source_span=_span(
                        project_id,
                        module_name,
                        "udt-field",
                        field.group(1),
                        source,
                        offsets,
                        start,
                        start + len(text),
                    ),
                )
            )
        index += 1
    if index >= len(lines):
        raise VBAParseError(f"unterminated Type {match.group(2)}")
    start = offsets[start_index]
    end = offsets[index] + len(lines[index].rstrip("\r\n"))
    return (
        VBAUserDefinedType(
            name=match.group(2),
            visibility=_visibility(match.group(1)),
            fields=fields,
            source_span=_span(
                project_id, module_name, "udt", match.group(2), source, offsets, start, end
            ),
        ),
        index + 1,
    )


def _variable(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    start: int,
    end: int,
    match: re.Match[str],
) -> VBAVariableDeclaration:
    qualifier = (match.group(1) or "Private").lower()
    visibility = _visibility(qualifier if qualifier in {"public", "private", "friend"} else None)
    bounds = []
    if match.group(4):
        bounds = [item.strip() for item in _split_commas(match.group(4)[1:-1])]
    return VBAVariableDeclaration(
        name=match.group(3),
        type_name=match.group(5) or "Variant",
        visibility=visibility,
        is_array=bool(match.group(4)),
        bounds=bounds,
        with_events=bool(match.group(2)),
        is_static=qualifier == "static",
        source_span=_span(
            project_id,
            module_name,
            "variable",
            f"{start}:{match.group(3)}",
            source,
            offsets,
            start,
            end,
        ),
    )


def _dependencies_for_line(
    project_id: str,
    module_name: str,
    source: str,
    offsets: list[int],
    start: int,
    end: int,
    line: str,
) -> list[VBAExternalDependency]:
    found: list[tuple[VBAExternalDependencyKind, str, str, dict[str, str]]] = []
    if match := _DECLARE.match(line):
        found.append(
            (
                VBAExternalDependencyKind.API,
                match.group(3),
                f"external_api:{match.group(3)}",
                {"procedure": match.group(2)},
            )
        )
    for match in re.finditer(
        r"\b(?:CreateObject|GetObject)\s*\(\s*\"([^\"]+)\"", line, re.IGNORECASE
    ):
        found.append((VBAExternalDependencyKind.COM, match.group(1), f"com:{match.group(1)}", {}))
    if re.search(r"\b(?:Open|Close|Kill|FileCopy|Dir|MkDir|RmDir)\b", line, re.IGNORECASE):
        found.append((VBAExternalDependencyKind.FILE, "vba-file-io", "filesystem", {}))
    if re.search(r"\b(?:WinHttp|XMLHTTP|InternetOpen|URLDownloadToFile)\b", line, re.IGNORECASE):
        found.append((VBAExternalDependencyKind.NETWORK, "network-access", "network", {}))
    if re.search(r"\b(?:ADODB|DAO\.|CurrentDb|OpenDatabase)\b", line, re.IGNORECASE):
        found.append((VBAExternalDependencyKind.DATABASE, "database-access", "database", {}))
    if re.search(r"\b(?:RegisterXLL|XLL)\b", line, re.IGNORECASE):
        found.append((VBAExternalDependencyKind.XLL, "xll-runtime", "xll", {}))
    result = []
    for position, (kind, dependency_name, capability, details) in enumerate(found):
        span = _span(
            project_id,
            module_name,
            "dependency",
            f"{start}:{position}:{dependency_name}",
            source,
            offsets,
            start,
            end,
        )
        result.append(
            VBAExternalDependency(
                dependency_id=span.node_id,
                kind=kind,
                name=dependency_name,
                capability=capability,
                details=details,
                source_span=span,
            )
        )
    return result


def _coerce_module_kind(
    name: str, source: str, explicit: VBAModuleKind | str | None
) -> VBAModuleKind:
    if explicit is not None:
        return explicit if isinstance(explicit, VBAModuleKind) else VBAModuleKind(explicit)
    lowered = name.lower().removesuffix(".bas").removesuffix(".cls").removesuffix(".frm")
    if lowered == "thisworkbook" or lowered.startswith("sheet"):
        return VBAModuleKind.DOCUMENT
    if name.lower().endswith(".frm") or lowered.startswith(("userform", "frm")):
        return VBAModuleKind.USERFORM
    if name.lower().endswith(".cls") or "Attribute VB_Exposed = True" in source:
        return VBAModuleKind.CLASS
    return VBAModuleKind.STANDARD


def _procedure_kind(value: str) -> VBAProcedureKind:
    normalized = re.sub(r"\s+", "_", value.strip().lower())
    return VBAProcedureKind(normalized)


def _visibility(value: str | None) -> VBAVisibility:
    return VBAVisibility((value or "private").lower())


def _statement_mutates_global(statement: VBAStatementIR) -> bool:
    return statement.kind in {VBAStatementKind.ASSIGNMENT, VBAStatementKind.REDIM} and bool(
        re.match(r"^(?:Public\s+)?[A-Za-z_]\w*\s*=", statement.text, re.IGNORECASE)
    )


def _called_names(line: str) -> set[str]:
    reserved = {"if", "for", "while", "do", "select", "redim", "array", "cells", "range"}
    return {
        match.group(1)
        for match in re.finditer(r"\b([A-Za-z_]\w*)\s*\(", line)
        if match.group(1).lower() not in reserved
    }


def _split_commas(text: str) -> list[str]:
    parts: list[str] = []
    current: list[str] = []
    depth = 0
    in_string = False
    index = 0
    while index < len(text):
        character = text[index]
        if character == '"':
            if in_string and index + 1 < len(text) and text[index + 1] == '"':
                current.extend(['"', '"'])
                index += 2
                continue
            in_string = not in_string
        elif not in_string and character == "(":
            depth += 1
        elif not in_string and character == ")":
            depth -= 1
        if character == "," and not in_string and depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(character)
        index += 1
    parts.append("".join(current))
    return parts


def _strip_comment(line: str) -> str:
    in_string = False
    index = 0
    while index < len(line):
        if line[index] == '"':
            if in_string and index + 1 < len(line) and line[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif line[index] == "'" and not in_string:
            return line[:index]
        index += 1
    return line


def _line_offsets(source: str) -> list[int]:
    offsets = [0]
    for match in re.finditer("\n", source):
        offsets.append(match.end())
    return offsets


def _line_column(offsets: list[int], offset: int) -> tuple[int, int]:
    line_index = 0
    for index, start in enumerate(offsets):
        if start > offset:
            break
        line_index = index
    return line_index + 1, offset - offsets[line_index] + 1


def _span(
    project_id: str,
    module_name: str,
    kind: str,
    name: str,
    source: str,
    offsets: list[int],
    start: int,
    end: int,
) -> SourceSpan:
    safe_end = max(start, min(end, len(source)))
    start_line, start_column = _line_column(offsets, start)
    end_line, end_column = _line_column(offsets, safe_end)
    identity = f"{project_id}\0{module_name}\0{kind}\0{name}\0{start}\0{safe_end}"
    return SourceSpan(
        node_id=f"vba-node:{_sha256(identity)[:24]}",
        module_name=module_name,
        start_offset=start,
        end_offset=safe_end,
        start_line=start_line,
        start_column=start_column,
        end_line=end_line,
        end_column=end_column,
        source_sha256=_sha256(source),
    )


def _source_map(project: VBAProjectIR) -> dict[str, SourceSpan]:
    spans: dict[str, SourceSpan] = {}

    def visit(value: Any) -> None:
        if isinstance(value, SourceSpan):
            spans[value.node_id] = value
        elif isinstance(value, BaseModel):
            for field_name in type(value).model_fields:
                if field_name != "source_map":
                    visit(getattr(value, field_name))
        elif isinstance(value, dict):
            for item in value.values():
                visit(item)
        elif isinstance(value, (list, tuple, set)):
            for item in value:
                visit(item)

    visit(project)
    return dict(sorted(spans.items()))


def _deduplicate_dependencies(
    dependencies: Any,
) -> list[VBAExternalDependency]:
    unique: dict[tuple[str, str, str], VBAExternalDependency] = {}
    for dependency in dependencies:
        key = (dependency.kind.value, dependency.name, dependency.capability)
        unique.setdefault(key, dependency)
    return [unique[key] for key in sorted(unique)]


def _sha256(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()
