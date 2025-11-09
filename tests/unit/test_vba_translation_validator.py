"""Unit tests for VBA translation validator."""

import pytest

from xlsliberator.vba_reference_analyzer import VBAReferences
from xlsliberator.vba_translation_validator import (
    TranslationEvaluation,
    TranslationIssue,
    TranslationIssueType,
    VBATranslationValidator,
)


@pytest.fixture
def validator() -> VBATranslationValidator:
    """Create validator instance."""
    return VBATranslationValidator()


@pytest.fixture
def simple_references() -> VBAReferences:
    """Create simple VBA references."""
    return VBAReferences(
        api_calls={"Range": 1},
        dependencies=set(),
        procedures=["Test"],
        special_patterns=[],
    )


def test_validator_initialization(validator: VBATranslationValidator) -> None:
    """Test validator initializes correctly."""
    assert validator.client is not None


def test_evaluate_valid_translation(
    validator: VBATranslationValidator, simple_references: VBAReferences
) -> None:
    """Test evaluating a valid translation."""
    vba_code = """
Sub Test()
    Range("A1").Value = "Hello"
End Sub
"""

    python_code = """
import uno
from loguru import logger

def Test():
    logger.info("Starting Test")
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    cell = sheet.getCellRangeByName("A1")
    cell.setString("Hello")
"""

    evaluation = validator.evaluate_translation(vba_code, python_code, simple_references)

    assert isinstance(evaluation, TranslationEvaluation)
    assert 0 <= evaluation.overall_quality <= 100
    assert isinstance(evaluation.is_acceptable, bool)
    assert isinstance(evaluation.issues, list)
    assert isinstance(evaluation.suggestions, list)


def test_evaluate_translation_with_issues(validator: VBATranslationValidator) -> None:
    """Test evaluating translation with obvious issues."""
    vba_code = """
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        Cells(i, 1).Value = i
    Next i
End Sub
"""

    # Bad translation with indexing error
    python_code = """
def Test():
    for i in range(1, 11):  # Wrong: should be range(10)
        cell = sheet.getCellByPosition(1, i)  # Wrong: 1-based instead of 0-based
        cell.setValue(i)
"""

    references = VBAReferences(
        api_calls={"Cells": 1},
        dependencies=set(),
        procedures=["Test"],
        special_patterns=["for_to_loop"],
    )

    evaluation = validator.evaluate_translation(vba_code, python_code, references)

    # Should detect issues (though quality might still be reasonable)
    assert evaluation.overall_quality < 100
    # May or may not have issues depending on LLM evaluation


def test_translation_issue_dataclass() -> None:
    """Test TranslationIssue dataclass."""
    issue = TranslationIssue(
        issue_type=TranslationIssueType.INDEXING_ERROR,
        description="1-based indexing used instead of 0-based",
        severity=8,
        vba_line="Cells(1, 1)",
        python_line="getCellByPosition(1, 1)",
    )

    assert issue.issue_type == TranslationIssueType.INDEXING_ERROR
    assert issue.severity == 8
    assert issue.vba_line == "Cells(1, 1)"
    assert issue.python_line == "getCellByPosition(1, 1)"


def test_translation_evaluation_dataclass() -> None:
    """Test TranslationEvaluation dataclass."""
    issue = TranslationIssue(
        issue_type=TranslationIssueType.MISSING_LOGIC,
        description="Loop body not translated",
        severity=9,
    )

    evaluation = TranslationEvaluation(
        overall_quality=65,
        is_acceptable=False,
        issues=[issue],
        suggestions=["Add loop translation", "Check variable scope"],
    )

    assert evaluation.overall_quality == 65
    assert evaluation.is_acceptable is False
    assert len(evaluation.issues) == 1
    assert len(evaluation.suggestions) == 2


def test_evaluation_conservative_fallback(validator: VBATranslationValidator) -> None:
    """Test that evaluation returns conservative result on failure."""
    # Test with empty code to potentially trigger errors
    vba_code = ""
    python_code = ""
    references = VBAReferences()

    # Should not raise, should return conservative evaluation
    evaluation = validator.evaluate_translation(vba_code, python_code, references)

    assert isinstance(evaluation, TranslationEvaluation)
    assert 0 <= evaluation.overall_quality <= 100


def test_format_api_usage(validator: VBATranslationValidator) -> None:
    """Test API usage formatting."""
    api_calls = {"Range": 10, "Cells": 5, "Worksheets": 2}
    formatted = validator._format_api_usage(api_calls)

    assert "Range: 10x" in formatted
    assert "Cells: 5x" in formatted
    assert "Worksheets: 2x" in formatted


def test_format_api_usage_empty(validator: VBATranslationValidator) -> None:
    """Test API usage formatting with empty input."""
    formatted = validator._format_api_usage({})
    assert "No API calls detected" in formatted


def test_format_patterns(validator: VBATranslationValidator) -> None:
    """Test pattern formatting."""
    patterns = ["for_each_loop", "error_handling", "arrays"]
    formatted = validator._format_patterns(patterns)

    assert "for_each_loop" in formatted
    assert "error_handling" in formatted
    assert "arrays" in formatted


def test_format_patterns_empty(validator: VBATranslationValidator) -> None:
    """Test pattern formatting with empty input."""
    formatted = validator._format_patterns([])
    assert "No special patterns detected" in formatted


def test_build_evaluation_prompt(
    validator: VBATranslationValidator, simple_references: VBAReferences
) -> None:
    """Test evaluation prompt building."""
    vba_code = 'Sub Test()\n    Range("A1").Value = "Hello"\nEnd Sub'
    python_code = (
        'def Test():\n    cell = sheet.getCellRangeByName("A1")\n    cell.setString("Hello")'
    )

    prompt = validator._build_evaluation_prompt(vba_code, python_code, simple_references)

    # Check key sections are present
    assert "Original VBA Code:" in prompt
    assert "Translated Python-UNO Code:" in prompt
    assert "Detected VBA API Usage:" in prompt
    assert "Detected VBA Patterns:" in prompt
    assert "overall_quality" in prompt
    assert "is_acceptable" in prompt
    assert "issues" in prompt


def test_all_issue_types_valid() -> None:
    """Test all TranslationIssueType enum values."""
    # Ensure all issue types can be constructed
    issue_types = [
        TranslationIssueType.SYNTAX_ERROR,
        TranslationIssueType.INCORRECT_API,
        TranslationIssueType.MISSING_LOGIC,
        TranslationIssueType.TYPE_MISMATCH,
        TranslationIssueType.INDEXING_ERROR,
        TranslationIssueType.ERROR_HANDLING,
        TranslationIssueType.CONTROL_FLOW,
        TranslationIssueType.VARIABLE_SCOPE,
    ]

    for issue_type in issue_types:
        issue = TranslationIssue(
            issue_type=issue_type,
            description=f"Test {issue_type.value}",
            severity=5,
        )
        assert issue.issue_type == issue_type
