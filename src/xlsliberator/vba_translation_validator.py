"""VBA translation validation using LLM reflection.

Validates VBA-to-Python-UNO translations by having Claude evaluate
its own translations and provide structured feedback for improvement.
"""

import json
import os
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import TYPE_CHECKING, Any

from anthropic import Anthropic
from loguru import logger

if TYPE_CHECKING:
    from xlsliberator.vba_reference_analyzer import VBAReferences


class TranslationIssueType(Enum):
    """Types of translation issues detected by LLM reflection."""

    SYNTAX_ERROR = "syntax_error"
    INCORRECT_API = "incorrect_api"
    MISSING_LOGIC = "missing_logic"
    TYPE_MISMATCH = "type_mismatch"
    INDEXING_ERROR = "indexing_error"  # 1-based vs 0-based
    ERROR_HANDLING = "error_handling"
    CONTROL_FLOW = "control_flow"
    VARIABLE_SCOPE = "variable_scope"


@dataclass
class TranslationIssue:
    """Identified issue in VBA-to-Python translation."""

    issue_type: TranslationIssueType
    description: str
    severity: int  # 1-10 (10 = critical)
    vba_line: str | None = None
    python_line: str | None = None


@dataclass
class TranslationEvaluation:
    """Evaluation result for a VBA-to-Python translation."""

    overall_quality: int  # 0-100
    is_acceptable: bool
    issues: list[TranslationIssue] = field(default_factory=list)
    suggestions: list[str] = field(default_factory=list)


class VBATranslationValidator:
    """Validates VBA-to-Python translations using LLM reflection."""

    def __init__(self) -> None:
        """Initialize translation validator."""
        self.client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    def evaluate_translation(
        self, vba_code: str, python_code: str, references: "VBAReferences"
    ) -> TranslationEvaluation:
        """Evaluate quality of VBA-to-Python translation using LLM.

        Args:
            vba_code: Original VBA source code
            python_code: Translated Python-UNO code
            references: VBA references (APIs, patterns)

        Returns:
            TranslationEvaluation with quality score, issues, and suggestions
        """
        logger.debug(
            f"Evaluating translation: VBA {len(vba_code)} bytes → Python {len(python_code)} bytes"
        )

        prompt = self._build_evaluation_prompt(vba_code, python_code, references)

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=4000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}],
            )

            # Parse structured evaluation response
            evaluation = self._parse_evaluation_response(response)

            logger.info(
                f"Translation evaluation: quality={evaluation.overall_quality}/100, "
                f"issues={len(evaluation.issues)}, acceptable={evaluation.is_acceptable}"
            )

            return evaluation

        except Exception as e:
            logger.error(f"Translation evaluation failed: {e}")
            # Return conservative evaluation on failure
            return TranslationEvaluation(
                overall_quality=50,
                is_acceptable=False,
                issues=[
                    TranslationIssue(
                        issue_type=TranslationIssueType.SYNTAX_ERROR,
                        description=f"Evaluation failed: {e}",
                        severity=5,
                    )
                ],
                suggestions=["Manual review recommended due to evaluation failure"],
            )

    def _build_evaluation_prompt(
        self, vba_code: str, python_code: str, references: "VBAReferences"
    ) -> str:
        """Build prompt for translation evaluation.

        Args:
            vba_code: Original VBA code
            python_code: Translated Python code
            references: VBA references

        Returns:
            Evaluation prompt string
        """
        api_usage_str = self._format_api_usage(references.api_calls)
        patterns_str = self._format_patterns(references.special_patterns)

        return f"""You are a VBA-to-Python-UNO translation validator.

Evaluate the quality of this VBA-to-Python translation.

Original VBA Code:
```vba
{vba_code}
```

Translated Python-UNO Code:
```python
{python_code}
```

Detected VBA API Usage:
{api_usage_str}

Detected VBA Patterns:
{patterns_str}

Your task:
1. Check if all VBA logic is preserved in Python
2. Verify all API calls are correctly translated (Range → sheet.getCellRangeByName, etc.)
3. Check for indexing errors (VBA 1-based → Python 0-based for getCellByPosition)
4. Verify error handling translation (On Error Resume Next → try/except)
5. Check variable declarations and types
6. Validate control flow (For Each, For To, Do While, Select Case)
7. Check that XSCRIPTCONTEXT is used for document access
8. Verify all required imports are present (uno, logger, etc.)

Provide evaluation in JSON format:
{{
    "overall_quality": <0-100>,
    "is_acceptable": <true if quality >= 80, else false>,
    "issues": [
        {{
            "type": "syntax_error|incorrect_api|missing_logic|type_mismatch|indexing_error|error_handling|control_flow|variable_scope",
            "severity": <1-10, where 10 is critical>,
            "description": "Clear description of the issue",
            "vba_line": "relevant VBA line or null",
            "python_line": "relevant Python line or null"
        }}
    ],
    "suggestions": [
        "Specific improvement suggestion 1",
        "Specific improvement suggestion 2"
    ]
}}

Important:
- Be strict about indexing: VBA Cells(1, 1) = (row 1, col 1) → getCellByPosition(0, 0)
- VBA Range("A1") → sheet.getCellRangeByName("A1")
- VBA Worksheets(1) → doc.getSheets().getByIndex(0)
- VBA MsgBox → logger.info() or print()
- VBA On Error Resume Next → try/except with logging

Respond with ONLY the JSON object, no explanations or markdown.
"""

    def _format_api_usage(self, api_calls: dict[str, int]) -> str:
        """Format API usage for prompt.

        Args:
            api_calls: Dictionary of API name -> count

        Returns:
            Formatted string
        """
        if not api_calls:
            return "(No API calls detected)"

        sorted_apis = sorted(api_calls.items(), key=lambda x: x[1], reverse=True)
        lines = [f"- {api}: {count}x" for api, count in sorted_apis[:10]]
        return "\n".join(lines)

    def _format_patterns(self, patterns: list[str]) -> str:
        """Format special patterns for prompt.

        Args:
            patterns: List of pattern identifiers

        Returns:
            Formatted string
        """
        if not patterns:
            return "(No special patterns detected)"

        return "\n".join(f"- {pattern}" for pattern in patterns)

    def _parse_evaluation_response(self, response: Any) -> TranslationEvaluation:
        """Parse Claude's evaluation response.

        Args:
            response: Anthropic API response

        Returns:
            TranslationEvaluation dataclass

        Raises:
            ValueError: If response format is unexpected
        """
        content_block = response.content[0]
        if hasattr(content_block, "text"):
            json_text = content_block.text.strip()
        else:
            raise ValueError(f"Unexpected content block type: {type(content_block)}")

        # Strip markdown code blocks if present
        json_text = re.sub(r"^```json\n|^```\n|```$", "", json_text, flags=re.MULTILINE)

        # Parse JSON
        try:
            evaluation_data: dict = json.loads(json_text)
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse evaluation JSON: {e}\nResponse: {json_text[:500]}")
            raise ValueError(f"Invalid JSON in evaluation response: {e}") from e

        # Convert to dataclass
        issues = [
            TranslationIssue(
                issue_type=TranslationIssueType(issue["type"]),
                description=issue["description"],
                severity=issue["severity"],
                vba_line=issue.get("vba_line"),
                python_line=issue.get("python_line"),
            )
            for issue in evaluation_data.get("issues", [])
        ]

        return TranslationEvaluation(
            overall_quality=evaluation_data["overall_quality"],
            is_acceptable=evaluation_data["is_acceptable"],
            issues=issues,
            suggestions=evaluation_data.get("suggestions", []),
        )
