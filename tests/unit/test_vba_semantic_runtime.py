"""Deterministic micro-conformance tests for VBA semantics and object model."""

import pytest

from xlsliberator.runtime.backend import FakeExcelBackend, RuntimeCapability
from xlsliberator.runtime.object_model import Application
from xlsliberator.vba_execution import (
    ByRefCell,
    TypedVBAInterpreter,
    VBAArray,
    VBAClassInstance,
    VBAEvent,
    VBAExecutionError,
    VBAVariant,
    coerce_number,
    coerce_string,
)
from xlsliberator.vba_parser import parse_vba_project


def test_vba_type_coercion_preserves_empty_null_boolean_and_error() -> None:
    assert coerce_number(VBAVariant.empty()) == 0
    assert coerce_string(VBAVariant.empty()) == ""
    assert coerce_number(VBAVariant.from_python(True)) == -1
    with pytest.raises(VBAExecutionError, match="Invalid use of Null"):
        coerce_number(VBAVariant.null())
    with pytest.raises(VBAExecutionError) as captured:
        coerce_number(VBAVariant.error(2023))
    assert captured.value.number == 2023


def test_byref_default_properties_arrays_and_redim_preserve() -> None:
    backend = FakeExcelBackend()
    application = Application(backend)
    target = application.active_sheet.range("A1")
    target.value = 7
    assert target.default == 7

    alias = ByRefCell(2)
    alias.value = 5
    assert alias.value == 5

    values = VBAArray(1, 2)
    values.set(1, "first")
    values.redim(1, 4, preserve=True)
    assert values.get(1).value == "first"
    assert values.get(4).kind.value == "empty"


def test_classes_events_and_range_collections_are_deterministic() -> None:
    instance = VBAClassInstance("Counter", "counter-1")
    instance.initialize()
    instance.fields["Value"] = VBAVariant.from_python(3)
    instance.terminate()
    assert instance.initialized and instance.terminated

    cancel = ByRefCell(False)
    event = VBAEvent("Workbook_BeforeClose", {"Cancel": cancel})
    event.cancel()
    assert event.cancelled and cancel.value is True

    backend = FakeExcelBackend()
    backend.names["Book1"] = {"Input": "Sheet1.A1"}
    backend.collections[("Book1", "Sheet1", RuntimeCapability.TABLES)] = [{"name": "T1"}]
    application = Application(backend)
    workbook = application.workbooks.item(1)
    assert workbook.names.item("Input").refers_to == "Sheet1.A1"
    assert workbook.worksheets.item(1).tables.items() == [{"name": "T1"}]
    assert workbook.emit_event("Workbook_Open")


def test_typed_interpreter_executes_byref_error_array_event_and_range_subset() -> None:
    project = parse_vba_project(
        "Micro",
        {
            "Module1": """Public Function Run(ByRef total As Long, Optional amount As Long = 2) As Variant
    Dim values() As Variant
    On Error Resume Next
    ReDim values(0 To 2)
    total = total + amount
    Range("A1").Value = total
    RaiseEvent Updated
    Run = total
End Function"""
        },
    )
    backend = FakeExcelBackend()
    interpreter = TypedVBAInterpreter(Application(backend))
    total = ByRefCell(3)

    result = interpreter.execute(project.modules[0].procedures[0], {"total": total})

    assert result.error_description is None
    assert total.value == 5
    assert result.return_value == 5
    assert backend.read_range("Book1", "Sheet1", "A1") == [[5]]
    assert result.events == ["Updated"]
    assert result.source_nodes_executed
