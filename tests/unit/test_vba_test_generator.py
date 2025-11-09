"""Unit tests for VBA test generator."""

import pytest

from xlsliberator.vba_reference_analyzer import VBAReferences
from xlsliberator.vba_test_generator import ValidationTest, VBATestGenerator


@pytest.fixture
def generator() -> VBATestGenerator:
    """Create test generator instance."""
    return VBATestGenerator()


@pytest.fixture
def simple_references() -> VBAReferences:
    """Create simple VBA references."""
    return VBAReferences(
        api_calls={"Range": 2, "Cells": 1},
        dependencies=set(),
        procedures=["TestMacro"],
        special_patterns=["for_to_loop"],
    )


def test_generator_initialization(generator: VBATestGenerator) -> None:
    """Test generator initializes correctly."""
    assert generator.client is not None


def test_validation_test_dataclass() -> None:
    """Test ValidationTest dataclass."""
    test = ValidationTest(
        test_name="test_cell_value",
        description="Verify cell A1 is set to 100",
        setup_code="sheet.getCellByPosition(0, 0).setValue(0)",
        vba_expected_behavior="Sets cell A1 to 100",
        python_test_code="SetCellValue()",
        assertions=["assert sheet.getCellByPosition(0, 0).getValue() == 100"],
    )

    assert test.test_name == "test_cell_value"
    assert test.description == "Verify cell A1 is set to 100"
    assert "setValue(0)" in test.setup_code
    assert "100" in test.vba_expected_behavior
    assert len(test.assertions) == 1


def test_generate_tests_simple_vba(
    generator: VBATestGenerator, simple_references: VBAReferences
) -> None:
    """Test generating tests for simple VBA code."""
    vba_code = """
Sub TestMacro()
    Range("A1").Value = 100
    Range("B1").Value = "Test"
End Sub
"""

    python_code = """
def TestMacro():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    sheet.getCellRangeByName("A1").setValue(100)
    sheet.getCellRangeByName("B1").setString("Test")
"""

    tests = generator.generate_tests(vba_code, python_code, simple_references, num_tests=2)

    assert isinstance(tests, list)
    assert len(tests) >= 1  # Should generate at least 1 test
    assert len(tests) <= 3  # Should not generate too many

    # Check first test structure
    test = tests[0]
    assert isinstance(test, ValidationTest)
    assert test.test_name != ""
    assert test.description != ""
    assert test.setup_code != ""
    assert test.vba_expected_behavior != ""
    assert test.python_test_code != ""
    assert len(test.assertions) > 0


def test_generate_tests_with_loop(generator: VBATestGenerator) -> None:
    """Test generating tests for VBA with loop."""
    vba_code = """
Sub FillRange()
    Dim i As Integer
    For i = 1 To 5
        Cells(i, 1).Value = i * 10
    Next i
End Sub
"""

    python_code = """
def FillRange():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    for i in range(5):
        sheet.getCellByPosition(0, i).setValue((i + 1) * 10)
"""

    references = VBAReferences(
        api_calls={"Cells": 5},
        dependencies=set(),
        procedures=["FillRange"],
        special_patterns=["for_to_loop"],
    )

    tests = generator.generate_tests(vba_code, python_code, references, num_tests=2)

    assert len(tests) >= 1
    # Should have assertions checking multiple cells
    test = tests[0]
    # At least one assertion should check cell values
    assert any("getCellByPosition" in assertion for assertion in test.assertions)


def test_format_api_usage(generator: VBATestGenerator) -> None:
    """Test API usage formatting."""
    api_calls = {"Range": 10, "Cells": 5, "Worksheets": 2}
    formatted = generator._format_api_usage(api_calls)

    assert "Range: 10x" in formatted
    assert "Cells: 5x" in formatted
    assert "Worksheets: 2x" in formatted


def test_format_api_usage_empty(generator: VBATestGenerator) -> None:
    """Test API usage formatting with empty input."""
    formatted = generator._format_api_usage({})
    assert "No API calls detected" in formatted


def test_format_patterns(generator: VBATestGenerator) -> None:
    """Test pattern formatting."""
    patterns = ["for_each_loop", "error_handling"]
    formatted = generator._format_patterns(patterns)

    assert "for_each_loop" in formatted
    assert "error_handling" in formatted


def test_format_patterns_empty(generator: VBATestGenerator) -> None:
    """Test pattern formatting with empty input."""
    formatted = generator._format_patterns([])
    assert "No special patterns detected" in formatted


def test_build_test_generation_prompt(
    generator: VBATestGenerator, simple_references: VBAReferences
) -> None:
    """Test test generation prompt building."""
    vba_code = 'Sub Test()\n    Range("A1").Value = 100\nEnd Sub'
    python_code = 'def Test():\n    sheet.getCellRangeByName("A1").setValue(100)'

    prompt = generator._build_test_generation_prompt(
        vba_code, python_code, simple_references, num_tests=3
    )

    # Check key sections are present
    assert "Original VBA Code:" in prompt
    assert "Translated Python-UNO Code:" in prompt
    assert "Detected API Usage:" in prompt
    assert "Detected Patterns:" in prompt
    assert "Generate 3" in prompt
    assert "test_name" in prompt
    assert "description" in prompt
    assert "setup_code" in prompt
    assert "assertions" in prompt


def test_generate_tests_with_complex_vba(generator: VBATestGenerator) -> None:
    """Test generating tests for complex VBA with multiple operations."""
    vba_code = """
Sub ProcessData()
    Dim sum As Double
    sum = 0

    For i = 1 To 10
        sum = sum + Cells(i, 1).Value
    Next i

    Range("B1").Value = sum
    Range("B2").Value = sum / 10
End Sub
"""

    python_code = """
def ProcessData():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    sum_val = 0.0

    for i in range(10):
        sum_val += sheet.getCellByPosition(0, i).getValue()

    sheet.getCellRangeByName("B1").setValue(sum_val)
    sheet.getCellRangeByName("B2").setValue(sum_val / 10)
"""

    references = VBAReferences(
        api_calls={"Cells": 10, "Range": 2},
        dependencies=set(),
        procedures=["ProcessData"],
        special_patterns=["for_to_loop"],
    )

    tests = generator.generate_tests(vba_code, python_code, references, num_tests=3)

    assert len(tests) >= 1
    # Should have setup code to populate input cells
    assert any("setup_code" in test.__dict__ and test.setup_code != "" for test in tests)
    # Should have assertions checking output cells
    assert any(len(test.assertions) > 0 for test in tests)


def test_generate_tests_error_handling(generator: VBATestGenerator) -> None:
    """Test that generator handles errors gracefully."""
    # Test with minimal/empty code
    vba_code = ""
    python_code = ""
    references = VBAReferences()

    # Should not crash, may return empty list or raise ValueError
    try:
        tests = generator.generate_tests(vba_code, python_code, references, num_tests=1)
        # If it succeeds, should return a list
        assert isinstance(tests, list)
    except ValueError:
        # Expected if LLM returns invalid JSON
        pass


def test_validation_test_all_fields_present() -> None:
    """Test that ValidationTest requires all fields."""
    # All fields should be present
    test = ValidationTest(
        test_name="test_name",
        description="description",
        setup_code="setup",
        vba_expected_behavior="expected",
        python_test_code="test_code",
        assertions=["assertion1"],
    )

    assert hasattr(test, "test_name")
    assert hasattr(test, "description")
    assert hasattr(test, "setup_code")
    assert hasattr(test, "vba_expected_behavior")
    assert hasattr(test, "python_test_code")
    assert hasattr(test, "assertions")
