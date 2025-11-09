"""Unit tests for VBA reference analyzer."""

from xlsliberator.vba_reference_analyzer import (
    VBAReferences,
    analyze_vba_references,
    get_recommended_translation_strategy,
    get_top_apis,
    get_translation_complexity_score,
)


def test_analyze_simple_vba() -> None:
    """Test analyzing simple VBA code."""
    vba_code = """
    Sub Test()
        Range("A1").Value = "Hello"
        MsgBox("Done")
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "Range" in refs.api_calls
    assert "MsgBox" in refs.api_calls
    assert len(refs.procedures) == 1
    assert "Test" in refs.procedures


def test_detect_error_handling() -> None:
    """Test detecting error handling patterns."""
    vba_code = """
    Sub Test()
        On Error Resume Next
        Range("A1").Value = 100
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "error_handling" in refs.special_patterns


def test_detect_for_each_loop() -> None:
    """Test detecting For Each loops."""
    vba_code = """
    Sub Test()
        Dim cell As Range
        For Each cell In Range("A1:A10")
            cell.Value = 100
        Next cell
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "for_each_loop" in refs.special_patterns
    assert "Range" in refs.api_calls


def test_detect_for_to_loop() -> None:
    """Test detecting For...To loops."""
    vba_code = """
    Sub Test()
        Dim i As Integer
        For i = 1 To 10
            Cells(i, 1).Value = i
        Next i
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "for_to_loop" in refs.special_patterns
    assert "Cells" in refs.api_calls


def test_detect_arrays() -> None:
    """Test detecting array declarations."""
    vba_code = """
    Sub Test()
        Dim arr(10) As Integer
        Dim i As Integer
        For i = 0 To 10
            arr(i) = i * 2
        Next i
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "arrays" in refs.special_patterns
    assert "for_to_loop" in refs.special_patterns


def test_detect_select_case() -> None:
    """Test detecting Select Case statements."""
    vba_code = """
    Sub Test()
        Dim x As Integer
        x = 5
        Select Case x
            Case 1
                MsgBox "One"
            Case 5
                MsgBox "Five"
        End Select
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "select_case" in refs.special_patterns


def test_detect_with_block() -> None:
    """Test detecting With blocks."""
    vba_code = """
    Sub Test()
        With Range("A1")
            .Value = 100
            .Font.Bold = True
        End With
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "with_block" in refs.special_patterns
    assert "Range" in refs.api_calls


def test_detect_worksheet_functions() -> None:
    """Test detecting WorksheetFunction usage."""
    vba_code = """
    Sub Test()
        Dim result As Double
        result = WorksheetFunction.Sum(Range("A1:A10"))
        MsgBox result
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "worksheet_functions" in refs.special_patterns
    assert "WorksheetFunction" in refs.api_calls
    assert "Range" in refs.api_calls


def test_detect_late_binding() -> None:
    """Test detecting late binding (CreateObject/GetObject)."""
    vba_code = """
    Sub Test()
        Dim dict As Object
        Set dict = CreateObject("Scripting.Dictionary")
        dict.Add "key", "value"
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "late_binding" in refs.special_patterns
    assert "CreateObject" in refs.api_calls
    assert "object_variables" in refs.special_patterns


def test_detect_exit_early() -> None:
    """Test detecting early exit statements."""
    vba_code = """
    Sub Test()
        If Range("A1").Value = "" Then
            Exit Sub
        End If
        MsgBox "Continuing"
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    assert "exit_early" in refs.special_patterns


def test_get_top_apis() -> None:
    """Test getting top APIs."""
    refs = VBAReferences(
        api_calls={"Range": 10, "Cells": 5, "Worksheets": 2, "MsgBox": 1},
        dependencies=set(),
        procedures=[],
        special_patterns=[],
    )

    top_apis = get_top_apis(refs, top_n=2)

    assert len(top_apis) == 2
    assert top_apis[0] == ("Range", 10)
    assert top_apis[1] == ("Cells", 5)


def test_complexity_score_simple() -> None:
    """Test complexity score for simple VBA."""
    refs = VBAReferences(
        api_calls={"Range": 1},
        dependencies=set(),
        procedures=["Test"],
        special_patterns=[],
    )

    score = get_translation_complexity_score(refs)

    # Simple code should have low complexity
    assert score <= 30


def test_complexity_score_moderate() -> None:
    """Test complexity score for moderate VBA."""
    refs = VBAReferences(
        api_calls={"Range": 5, "Cells": 3, "Worksheets": 2},
        dependencies={"ModuleB"},
        procedures=["Test", "Helper", "Validate"],
        special_patterns=["for_each_loop", "error_handling", "arrays"],
    )

    score = get_translation_complexity_score(refs)

    # Moderate complexity
    assert 31 <= score <= 60


def test_complexity_score_complex() -> None:
    """Test complexity score for complex VBA."""
    refs = VBAReferences(
        api_calls={
            "Range": 10,
            "Cells": 8,
            "Worksheets": 5,
            "WorksheetFunction": 3,
            "CreateObject": 2,
        },
        dependencies={"ModuleB", "Helper"},
        procedures=["Main", "Process", "Validate", "Format"],
        special_patterns=[
            "error_handling",
            "for_each_loop",
            "arrays",
            "select_case",
            "late_binding",
        ],
    )

    score = get_translation_complexity_score(refs)

    # Complex code should have high complexity
    assert score >= 61


def test_recommended_strategy_simple() -> None:
    """Test recommended strategy for simple VBA."""
    refs = VBAReferences(
        api_calls={"Range": 1},
        dependencies=set(),
        procedures=["Test"],
        special_patterns=[],
    )

    strategy = get_recommended_translation_strategy(refs)

    assert strategy == "rule_based"


def test_recommended_strategy_moderate() -> None:
    """Test recommended strategy for moderate VBA."""
    refs = VBAReferences(
        api_calls={"Range": 5, "Cells": 3, "Worksheets": 2},
        dependencies={"ModuleB"},
        procedures=["Test", "Helper", "Validate"],
        special_patterns=["for_each_loop", "error_handling", "arrays"],
    )

    strategy = get_recommended_translation_strategy(refs)

    assert strategy == "llm_basic"


def test_recommended_strategy_complex() -> None:
    """Test recommended strategy for complex VBA."""
    refs = VBAReferences(
        api_calls={
            "Range": 10,
            "Cells": 8,
            "Worksheets": 5,
            "WorksheetFunction": 3,
            "CreateObject": 2,
            "Application": 4,
        },
        dependencies={"ModuleB", "Helper", "Utils"},
        procedures=["Main", "Process", "Validate", "Format", "Export"],
        special_patterns=[
            "error_handling",
            "for_each_loop",
            "arrays",
            "select_case",
            "late_binding",
            "with_block",
            "property_procedures",
        ],
    )

    strategy = get_recommended_translation_strategy(refs)

    assert strategy == "llm_reflection"


def test_multiple_patterns() -> None:
    """Test detecting multiple patterns in complex VBA."""
    vba_code = """
    Sub ComplexTest()
        On Error Resume Next
        Dim arr(10) As Integer
        Dim i As Integer
        Dim cell As Range

        For i = 1 To 10
            arr(i) = i * 2
        Next i

        For Each cell In Range("A1:A10")
            cell.Value = arr(cell.Row)
        Next cell

        Dim result As Double
        result = WorksheetFunction.Sum(Range("A1:A10"))

        Select Case result
            Case Is > 100
                MsgBox("High")
            Case Else
                MsgBox("Low")
        End Select

        If result > 50 Then
            Exit Sub
        End If
    End Sub
    """

    refs = analyze_vba_references(vba_code)

    # Should detect multiple patterns
    assert "error_handling" in refs.special_patterns
    assert "arrays" in refs.special_patterns
    assert "for_to_loop" in refs.special_patterns
    assert "for_each_loop" in refs.special_patterns
    assert "worksheet_functions" in refs.special_patterns
    assert "select_case" in refs.special_patterns
    assert "exit_early" in refs.special_patterns

    # Should detect multiple APIs
    assert "Range" in refs.api_calls
    assert "WorksheetFunction" in refs.api_calls
    assert "MsgBox" in refs.api_calls

    # Should have moderate to high complexity (7 patterns detected)
    complexity = get_translation_complexity_score(refs)
    assert complexity >= 40

    # Should recommend at least llm_basic strategy
    strategy = get_recommended_translation_strategy(refs)
    assert strategy in ["llm_basic", "llm_reflection"]
