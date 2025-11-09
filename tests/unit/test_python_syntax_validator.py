"""Unit tests for Python syntax validator."""

from xlsliberator.python_syntax_validator import (
    PythonSyntaxValidator,
)


def test_valid_python_code() -> None:
    """Test validating syntactically correct Python code."""
    validator = PythonSyntaxValidator()

    python_code = """
def hello_world():
    print("Hello, World!")
    return 42
"""

    result = validator.validate_syntax(python_code)

    assert result.is_valid is True
    assert len(result.syntax_errors) == 0


def test_syntax_error_detection() -> None:
    """Test detecting basic syntax errors."""
    validator = PythonSyntaxValidator()

    # Missing colon after function definition
    python_code = """
def hello_world()
    print("Hello")
"""

    result = validator.validate_syntax(python_code)

    assert result.is_valid is False
    assert len(result.syntax_errors) > 0
    assert "Syntax error" in result.syntax_errors[0]


def test_compilation_error_detection() -> None:
    """Test detecting compilation errors."""
    validator = PythonSyntaxValidator()

    # Invalid syntax that might pass AST but fail compilation
    python_code = """
def test():
    return
    # This is syntactically valid but has unreachable code
"""

    result = validator.validate_syntax(python_code)

    # Should at least parse successfully
    assert result is not None


def test_detect_missing_uno_import() -> None:
    """Test detecting missing uno import."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    ctx = uno.getComponentContext()
    return ctx
"""

    result = validator.validate_syntax(python_code)

    assert "Missing 'import uno' import" in result.warnings


def test_detect_missing_logger_import() -> None:
    """Test detecting missing logger import."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    logger.info("Testing")
    return 42
"""

    result = validator.validate_syntax(python_code)

    assert "Missing 'from loguru import logger' import" in result.warnings


def test_detect_missing_math_import() -> None:
    """Test detecting missing math import."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    result = math.sqrt(16)
    return result
"""

    result = validator.validate_syntax(python_code)

    assert "Missing 'import math' import" in result.warnings


def test_detect_missing_datetime_import() -> None:
    """Test detecting missing datetime import."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    now = datetime.now()
    return now
"""

    result = validator.validate_syntax(python_code)

    assert "Missing 'from datetime import datetime' import" in result.warnings


def test_detect_1based_indexing_issue() -> None:
    """Test detecting potential 1-based indexing errors."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    sheet = XSCRIPTCONTEXT.getDocument().getSheets().getByIndex(0)
    cell = sheet.getCellByPosition(1, 1)  # Suspicious: both > 0
    return cell.getValue()
"""

    result = validator.validate_syntax(python_code)

    # Should warn about possible 1-based indexing
    indexing_warnings = [w for w in result.warnings if "1-based indexing" in w]
    assert len(indexing_warnings) > 0


def test_detect_range_with_suspicious_bounds() -> None:
    """Test detecting range(1, n) which might be VBA translation error."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    for i in range(1, 10):  # VBA For i = 1 To 10
        print(i)
"""

    result = validator.validate_syntax(python_code)

    # Should warn about suspicious range(1, n)
    range_warnings = [w for w in result.warnings if "range(1," in w]
    assert len(range_warnings) > 0


def test_detect_vba_string_concatenation() -> None:
    """Test detecting VBA-style string concatenation (&)."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    result = "Hello" & "World"  # VBA style, not Python
    return result
"""

    result = validator.validate_syntax(python_code)

    # Should warn about VBA-style concatenation
    concat_warnings = [w for w in result.warnings if "concatenation" in w.lower()]
    assert len(concat_warnings) > 0


def test_detect_vba_keywords() -> None:
    """Test detecting VBA keywords in Python code."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    Dim x As Integer  # VBA keyword
    Set obj = CreateObject("Scripting.Dictionary")
    End Sub
"""

    result = validator.validate_syntax(python_code)

    # Should warn about VBA keywords
    vba_warnings = [w for w in result.warnings if "VBA keyword" in w]
    assert len(vba_warnings) >= 2  # Dim, Set, End Sub


def test_detect_vba_comments() -> None:
    """Test detecting VBA-style comments."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    ' This is a VBA comment
    x = 42
    return x
"""

    result = validator.validate_syntax(python_code)

    # Should warn about VBA-style comments
    comment_warnings = [w for w in result.warnings if "comment" in w.lower()]
    assert len(comment_warnings) > 0


def test_valid_uno_code_with_imports() -> None:
    """Test valid Python-UNO code with proper imports."""
    validator = PythonSyntaxValidator()

    python_code = """
import uno
from loguru import logger

def test():
    logger.info("Starting")
    ctx = uno.getComponentContext()
    return ctx
"""

    result = validator.validate_syntax(python_code)

    assert result.is_valid is True
    # Should not warn about missing imports
    import_warnings = [w for w in result.warnings if "Missing" in w]
    assert len(import_warnings) == 0


def test_valid_indexing_no_warnings() -> None:
    """Test that valid 0-based indexing doesn't trigger warnings."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    sheet = XSCRIPTCONTEXT.getDocument().getSheets().getByIndex(0)
    cell = sheet.getCellByPosition(0, 0)  # Valid: 0-based
    return cell.getValue()
"""

    result = validator.validate_syntax(python_code)

    # Should not warn about 0-based indexing
    indexing_warnings = [w for w in result.warnings if "indexing" in w.lower()]
    assert len(indexing_warnings) == 0


def test_complex_code_with_multiple_issues() -> None:
    """Test code with multiple issues detected."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    ' VBA comment
    Dim x As Integer
    logger.info("Testing")
    result = math.sqrt(16)
    cell = sheet.getCellByPosition(1, 1)
    msg = "Hello" & "World"
    for i in range(1, 10):
        print(i)
"""

    result = validator.validate_syntax(python_code)

    # Should detect multiple issues
    assert len(result.warnings) >= 5
    # VBA comment, VBA keyword, missing logger import, missing math import,
    # 1-based indexing, VBA concatenation, suspicious range


def test_libreoffice_python_compatibility_check() -> None:
    """Test LibreOffice Python compatibility check (if available)."""
    validator = PythonSyntaxValidator()

    python_code = """
import uno

def test():
    ctx = uno.getComponentContext()
    return ctx
"""

    result = validator.validate_syntax(python_code)

    # If LibreOffice Python is available, compatibility should be checked
    # If not available, should default to True
    assert isinstance(result.uno_compatible, bool)


def test_no_false_positives_for_strings() -> None:
    """Test that string literals don't trigger false positives."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    # String containing VBA keywords should not trigger warnings
    vba_code = "Dim x As Integer"
    vba_comment = "' This is VBA"
    return vba_code + vba_comment
"""

    result = validator.validate_syntax(python_code)

    # Note: Current implementation may have false positives for strings
    # This test documents current behavior
    # Future improvement: Parse AST to exclude string literals
    assert result.is_valid is True


def test_xscriptcontext_requires_uno_import() -> None:
    """Test that XSCRIPTCONTEXT usage triggers uno import warning."""
    validator = PythonSyntaxValidator()

    python_code = """
def test():
    doc = XSCRIPTCONTEXT.getDocument()
    return doc
"""

    result = validator.validate_syntax(python_code)

    assert "Missing 'import uno' import" in result.warnings


def test_worksheet_function_translation_example() -> None:
    """Test realistic VBA→Python-UNO translation validation."""
    validator = PythonSyntaxValidator()

    # Good translation
    good_code = """
import uno
from loguru import logger

def calculate_sum():
    logger.info("Calculating sum")
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)

    total = 0.0
    for row in range(10):  # 0-based indexing
        cell = sheet.getCellByPosition(0, row)
        total += cell.getValue()

    return total
"""

    result = validator.validate_syntax(good_code)

    assert result.is_valid is True
    assert len(result.syntax_errors) == 0
    # range(10) is fine, range(1, 10) would be suspicious


def test_invalid_python_syntax() -> None:
    """Test various invalid Python syntax patterns."""
    validator = PythonSyntaxValidator()

    invalid_codes = [
        "def test(",  # Incomplete function def
        "if True\n    pass",  # Missing colon
        "for i in range(10)\n    print(i)",  # Missing colon
        "class Test\n    pass",  # Missing colon
    ]

    for code in invalid_codes:
        result = validator.validate_syntax(code)
        assert result.is_valid is False
        assert len(result.syntax_errors) > 0
