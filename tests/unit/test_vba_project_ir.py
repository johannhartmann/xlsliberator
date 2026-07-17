"""Typed VBA project parser and fail-closed execution-plan tests."""

from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    ExternalCapability,
    ExternalCapabilityKind,
)
from xlsliberator.vba_execution import (
    DifferentialProof,
    ProcedureStrategy,
    build_execution_plan,
)
from xlsliberator.vba_ir import (
    VBAExternalDependencyKind,
    VBAModuleKind,
    VBAParameterPassing,
    VBAProcedureKind,
    VBAProjectReference,
    VBAStatementKind,
)
from xlsliberator.vba_parser import parse_vba_project

MODULE = """Attribute VB_Name = "Module1"
#Const Win64 = True
#If Win64 Then
Private Declare PtrSafe Function Tick Lib "kernel32" () As Long
#Else
Private Declare Function Tick Lib "kernel32" () As Long
#End If
Public globalValue As Variant
Private Const Limit As Long = 4
Public Enum Mode
    ModeOne = 1
End Enum
Public Type Point
    X As Double
    Y As Double
End Type

Public Function Compute(ByRef total As Long, ByVal value As Double, Optional label As String = "x", ParamArray rest() As Variant) As Variant
    Dim values() As Variant
    On Error Resume Next
    ReDim values(0 To 2)
    total = total + value
    Range("A1").Value = total
    If value > 0 Then
        CreateObject("Scripting.Dictionary")
    End If
    Compute = total
End Function
"""

EVENT_MODULE = """Attribute VB_Name = "ThisWorkbook"
Private WithEvents app As Application
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    RaiseEvent Closing
    Cancel = True
End Sub
"""


def _project():  # type: ignore[no-untyped-def]
    return parse_vba_project(
        "Conformance",
        {"Module1.bas": MODULE, "ThisWorkbook.cls": EVENT_MODULE},
        module_kinds={"ThisWorkbook.cls": VBAModuleKind.DOCUMENT},
        references=[
            VBAProjectReference(
                name="Scripting",
                guid="{420B2830-E718-11CF-893D-00A0C9054228}",
                major=1,
                minor=0,
                capability="project_reference:Scripting",
            )
        ],
        conditional_compilation_arguments={"Win64": "True"},
    )


def test_parser_represents_project_semantics_and_stable_source_map() -> None:
    first = _project()
    second = _project()

    assert first.model_dump(mode="json") == second.model_dump(mode="json")
    assert first.schema_version == "1.0.0"
    assert first.references[0].guid
    assert first.conditional_compilation_arguments == {"Win64": "True"}
    module = next(module for module in first.modules if module.name == "Module1.bas")
    assert module.kind is VBAModuleKind.STANDARD
    assert module.conditional_constants[0].name == "Win64"
    assert module.conditional_blocks[0].branches == ["Win64", "else"]
    assert module.variables[0].name == "globalValue"
    assert module.constants[0].name == "Limit"
    assert module.enums[0].members[0].name == "ModeOne"
    assert [field.name for field in module.user_defined_types[0].fields] == ["X", "Y"]
    procedure = module.procedures[0]
    assert procedure.kind is VBAProcedureKind.FUNCTION
    assert [parameter.passing for parameter in procedure.parameters[:2]] == [
        VBAParameterPassing.BYREF,
        VBAParameterPassing.BYVAL,
    ]
    assert procedure.parameters[2].optional
    assert procedure.parameters[3].param_array
    assert procedure.parameters[3].is_array
    assert VBAStatementKind.ON_ERROR in {statement.kind for statement in procedure.statements}
    assert VBAStatementKind.REDIM in {statement.kind for statement in procedure.statements}
    assert VBAStatementKind.IF in {statement.kind for statement in procedure.statements}
    assert any(
        expression.uses_default_member
        for statement in procedure.statements
        for expression in statement.expressions
    )
    assert any(
        dependency.kind is VBAExternalDependencyKind.COM
        for dependency in procedure.external_dependencies
    )
    assert procedure.source_span.node_id in first.source_map
    assert all(node_id == span.node_id for node_id, span in first.source_map.items())


def test_parser_represents_events_cancellation_lifetime_and_global_state() -> None:
    project = _project()
    module = next(module for module in project.modules if module.kind is VBAModuleKind.DOCUMENT)
    procedure = module.procedures[0]

    assert module.lifetime == "document"
    assert module.has_global_state
    assert module.variables[0].with_events
    assert procedure.is_event_handler
    assert procedure.cancels_event
    assert VBAStatementKind.RAISE_EVENT in {statement.kind for statement in procedure.statements}


def test_external_dependencies_are_typed_capabilities_and_block_plan() -> None:
    project = _project()
    plan = build_execution_plan(project, EnvironmentManifest())

    assert not plan.fully_executable
    assert "com:Scripting.Dictionary" in plan.missing_capabilities
    assert "external_api:kernel32" in plan.missing_capabilities
    assert "project_reference:Scripting" in plan.missing_capabilities
    assert any(not decision.executable for decision in plan.decisions)


def test_python_translation_requires_differential_proof() -> None:
    project = parse_vba_project("Simple", {"Module1": "Public Sub Run()\n    x = 1\nEnd Sub"})
    procedure = project.modules[0].procedures[0]
    environment = EnvironmentManifest(
        typed_capabilities=[
            ExternalCapability(
                capability="vba.python_runtime",
                kind=ExternalCapabilityKind.ADD_IN,
                resource="xlsliberator-runtime",
                granted=True,
            )
        ]
    )

    missing = build_execution_plan(
        project,
        environment,
        preferred_strategy=ProcedureStrategy.TRANSLATE_PYTHON,
    )
    proven = build_execution_plan(
        project,
        environment,
        preferred_strategy=ProcedureStrategy.TRANSLATE_PYTHON,
        differential_proofs=[
            DifferentialProof(
                procedure_id=procedure.procedure_id,
                source_trace_id="excel-trace",
                target_trace_id="libreoffice-trace",
                equivalent=True,
            )
        ],
    )

    assert not missing.fully_executable
    assert proven.fully_executable
    assert proven.decisions[0].strategy is ProcedureStrategy.TRANSLATE_PYTHON
