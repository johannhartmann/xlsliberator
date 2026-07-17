"""Python-UNO syntax validation for translated VBA code.

Validates generated Python code for:
- Basic syntax errors (AST parsing)
- Compilation errors (py_compile)
- Common translation mistakes (indexing, imports)
- LibreOffice Python compatibility (optional)
"""

import ast
import py_compile
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from loguru import logger


@dataclass
class SyntaxValidationResult:
    """Result of Python syntax validation."""

    is_valid: bool
    syntax_errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    uno_compatible: bool = True


class PythonSyntaxValidator:
    """Validates Python-UNO code syntax."""

    def __init__(self, libreoffice_python_path: Path | None = None):
        """Initialize syntax validator.

        Args:
            libreoffice_python_path: Path to LibreOffice's Python interpreter
                                    (e.g., /usr/lib/libreoffice/program/python)
                                    If None, will attempt auto-detection.
        """
        # A legacy path may request a runtime check, but can never select a host
        # interpreter. Compatibility checks are routed to the Docker worker.
        self.check_runtime_compatibility = libreoffice_python_path is not None
        self.libreoffice_python = None

    def _find_libreoffice_python(self) -> Path | None:
        """Auto-detect LibreOffice Python interpreter.

        Returns:
            Path to LibreOffice Python, or None if not found
        """
        logger.debug("Host LibreOffice Python discovery is disabled")
        return None

    def validate_syntax(self, python_code: str) -> SyntaxValidationResult:
        """Validate Python syntax and detect common issues.

        Args:
            python_code: Python source code to validate

        Returns:
            SyntaxValidationResult with validation details
        """
        errors: list[str] = []
        warnings: list[str] = []

        # 1. AST parsing check (basic syntax)
        try:
            ast.parse(python_code)
            logger.debug("AST parsing succeeded")
        except SyntaxError as e:
            error_msg = f"Syntax error at line {e.lineno}: {e.msg}"
            if e.text:
                error_msg += f" in '{e.text.strip()}'"
            errors.append(error_msg)
            logger.error(error_msg)

        # 2. Compile check (more thorough than AST)
        if not errors:  # Only if AST passed
            compile_errors = self._check_compilation(python_code)
            errors.extend(compile_errors)

        # 3. Static analysis for common translation issues
        static_warnings = self._analyze_common_issues(python_code)
        warnings.extend(static_warnings)

        # 4. LibreOffice Python compatibility (if available)
        uno_compatible = True
        if self.check_runtime_compatibility:
            uno_compatible = self._check_uno_compatibility(python_code)
            if not uno_compatible:
                warnings.append("UNO compatibility check failed - code may not work in LibreOffice")

        return SyntaxValidationResult(
            is_valid=len(errors) == 0,
            syntax_errors=errors,
            warnings=warnings,
            uno_compatible=uno_compatible,
        )

    def _check_compilation(self, python_code: str) -> list[str]:
        """Check if code compiles using py_compile.

        Args:
            python_code: Python source code

        Returns:
            List of compilation errors (empty if successful)
        """
        errors = []

        with tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False) as f:
            f.write(python_code)
            temp_path = f.name

        try:
            py_compile.compile(temp_path, doraise=True)
            logger.debug("py_compile succeeded")
        except py_compile.PyCompileError as e:
            errors.append(f"Compilation error: {e.msg}")
            logger.error(f"Compilation failed: {e}")
        finally:
            Path(temp_path).unlink(missing_ok=True)

        return errors

    def _check_uno_compatibility(self, python_code: str) -> bool:
        """Check if code is compatible with LibreOffice Python.

        Args:
            python_code: Python source code

        Returns:
            True if compatible, False otherwise
        """
        from xlsliberator.lo_worker_client import LibreOfficeWorkerClient

        response = LibreOfficeWorkerClient(timeout_seconds=10).request(
            {"op": "validate_python", "python_code": python_code},
            timeout_seconds=10,
        )
        if not response.success:
            logger.warning("Docker UNO compatibility check unavailable or failed")
            return False
        return bool(response.data.get("compatible"))

    def _analyze_common_issues(self, python_code: str) -> list[str]:
        """Analyze code for common VBA translation issues.

        Args:
            python_code: Python source code

        Returns:
            List of warning messages
        """
        warnings = []

        # Check for possible 1-based indexing errors
        indexing_warnings = self._check_indexing_issues(python_code)
        warnings.extend(indexing_warnings)

        # Check for missing imports
        import_warnings = self._check_missing_imports(python_code)
        warnings.extend(import_warnings)

        # Check for VBA-style patterns that shouldn't be in Python
        vba_pattern_warnings = self._check_vba_patterns(python_code)
        warnings.extend(vba_pattern_warnings)

        return warnings

    def _check_indexing_issues(self, python_code: str) -> list[str]:
        """Check for potential 1-based indexing errors.

        Args:
            python_code: Python source code

        Returns:
            List of warnings about indexing
        """
        warnings = []

        # Check getCellByPosition calls with positive integers
        # UNO uses 0-based indexing, VBA uses 1-based
        pattern = r"getCellByPosition\((\d+)\s*,\s*(\d+)\)"
        matches = re.finditer(pattern, python_code)

        for match in matches:
            col_str, row_str = match.group(1), match.group(2)
            col, row = int(col_str), int(row_str)

            # If both indices are > 0, might be 1-based indexing
            if col > 0 and row > 0:
                warnings.append(
                    f"Possible 1-based indexing in getCellByPosition({col}, {row}) "
                    f"at position {match.start()}. UNO uses 0-based indexing. "
                    f"VBA Cells({row + 1}, {col + 1}) → getCellByPosition({col}, {row})"
                )

        # Check for range() with suspicious bounds
        # for i in range(1, 10) is common in VBA translation but often wrong
        pattern = r"range\(1\s*,\s*(\d+)\)"
        matches = re.finditer(pattern, python_code)

        for match in matches:
            warnings.append(
                f"Suspicious range(1, {match.group(1)}) at position {match.start()}. "
                "VBA For i = 1 To n uses 1-based indexing. "
                "Python range(1, n+1) includes 1 but excludes n+1."
            )

        return warnings

    def _check_missing_imports(self, python_code: str) -> list[str]:
        """Check for missing required imports.

        Args:
            python_code: Python source code

        Returns:
            List of warnings about missing imports
        """
        warnings = []

        # Check if logger is used but not imported
        if "logger." in python_code and "from loguru import logger" not in python_code:
            warnings.append("Missing 'from loguru import logger' import")

        # Check if uno is used but not imported
        if (
            any(pattern in python_code for pattern in ["uno.", "XSCRIPTCONTEXT"])
            and "import uno" not in python_code
        ):
            warnings.append("Missing 'import uno' import")

        # Check if math functions are used but math not imported
        math_funcs = ["math.sqrt", "math.floor", "math.ceil", "math.pow"]
        if any(func in python_code for func in math_funcs) and "import math" not in python_code:
            warnings.append("Missing 'import math' import")

        # Check if datetime is used but not imported
        if "datetime." in python_code and "from datetime import datetime" not in python_code:
            warnings.append("Missing 'from datetime import datetime' import")

        return warnings

    def _check_vba_patterns(self, python_code: str) -> list[str]:
        """Check for VBA-style patterns that shouldn't appear in Python.

        Args:
            python_code: Python source code

        Returns:
            List of warnings about VBA patterns
        """
        warnings = []

        # Check for VBA-style string concatenation (&)
        # This should be + or f-strings in Python
        if " & " in python_code or '"&' in python_code or '&"' in python_code:
            warnings.append(
                "VBA-style string concatenation '&' found. Use '+' or f-strings in Python."
            )

        # Check for VBA keywords that shouldn't appear
        vba_keywords = [
            "Dim ",
            "Set ",
            "End Sub",
            "End Function",
            "End If",
            "Next ",
            "Loop",
        ]

        for keyword in vba_keywords:
            if keyword in python_code:
                warnings.append(f"VBA keyword '{keyword.strip()}' found in Python code")

        # Check for VBA-style comments (')
        if re.search(r"^\s*'", python_code, re.MULTILINE):
            warnings.append("VBA-style comment (') found. Use '#' for Python comments.")

        return warnings
