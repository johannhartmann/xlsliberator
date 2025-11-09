"""VBA validation test generator using LLM.

Generates validation tests that verify translated Python-UNO code
produces the same behavior as the original VBA code.
"""

import json
import os
import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from anthropic import Anthropic
from loguru import logger

if TYPE_CHECKING:
    from xlsliberator.vba_reference_analyzer import VBAReferences


@dataclass
class ValidationTest:
    """Generated validation test for VBA translation."""

    test_name: str
    description: str
    setup_code: str  # Set up test environment (e.g., populate cells)
    vba_expected_behavior: str  # What VBA would do
    python_test_code: str  # Python code that runs the translated function
    assertions: list[str]  # Expected assertions to validate behavior


class VBATestGenerator:
    """Generates validation tests for VBA-to-Python translations."""

    def __init__(self) -> None:
        """Initialize test generator."""
        self.client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    def generate_tests(
        self,
        vba_code: str,
        python_code: str,
        references: "VBAReferences",
        num_tests: int = 3,
    ) -> list[ValidationTest]:
        """Generate validation tests for VBA translation.

        Args:
            vba_code: Original VBA code
            python_code: Translated Python-UNO code
            references: Detected VBA references (APIs, patterns)
            num_tests: Number of tests to generate (default: 3)

        Returns:
            List of ValidationTest objects

        Raises:
            ValueError: If test generation fails or returns invalid JSON
        """
        logger.debug(f"Generating {num_tests} validation tests for VBA translation")

        prompt = self._build_test_generation_prompt(vba_code, python_code, references, num_tests)

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=8000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}],
            )

            content_block = response.content[0]
            if hasattr(content_block, "text"):
                tests_json = content_block.text.strip()
            else:
                raise ValueError(f"Unexpected content block type: {type(content_block)}")

            # Strip markdown code blocks if present
            tests_json = re.sub(r"^```json\n|^```\n|```$", "", tests_json, flags=re.MULTILINE)

            # Parse JSON
            tests_data: dict = json.loads(tests_json)

            # Convert to dataclasses
            tests = [
                ValidationTest(
                    test_name=test["test_name"],
                    description=test["description"],
                    setup_code=test["setup_code"],
                    vba_expected_behavior=test["vba_expected_behavior"],
                    python_test_code=test["python_test_code"],
                    assertions=test["assertions"],
                )
                for test in tests_data.get("tests", [])
            ]

            logger.info(f"Generated {len(tests)} validation tests successfully")
            return tests

        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse test generation JSON: {e}")
            raise ValueError(f"Invalid JSON in test generation response: {e}") from e
        except KeyError as e:
            logger.error(f"Missing required field in test generation response: {e}")
            raise ValueError(f"Missing required field in test data: {e}") from e
        except Exception as e:
            logger.error(f"Test generation failed: {e}")
            raise

    def _build_test_generation_prompt(
        self,
        vba_code: str,
        python_code: str,
        references: "VBAReferences",
        num_tests: int,
    ) -> str:
        """Build prompt for LLM test generation.

        Args:
            vba_code: Original VBA code
            python_code: Translated Python code
            references: VBA references
            num_tests: Number of tests to generate

        Returns:
            Test generation prompt string
        """
        api_usage_str = self._format_api_usage(references.api_calls)
        patterns_str = self._format_patterns(references.special_patterns)

        return f"""You are a VBA testing expert. Generate validation tests for this VBA-to-Python-UNO translation.

Original VBA Code:
```vba
{vba_code}
```

Translated Python-UNO Code:
```python
{python_code}
```

Detected API Usage:
{api_usage_str}

Detected Patterns:
{patterns_str}

Your task:
Generate {num_tests} validation tests that verify the Python translation produces the same behavior as the VBA code.

For each test:
1. **setup_code**: Python code to prepare the test environment
   - Create a mock XSCRIPTCONTEXT if needed
   - Set initial cell values
   - Prepare any required data structures

2. **vba_expected_behavior**: Describe what the VBA code would do
   - Be specific about cell values, formatting, etc.

3. **python_test_code**: Python code that calls the translated function
   - Use the same inputs as VBA would receive

4. **assertions**: List of Python assert statements to verify behavior
   - Check cell values: assert sheet.getCellByPosition(col, row).getValue() == expected
   - Check strings: assert sheet.getCellByPosition(col, row).getString() == "expected"
   - Check formulas if applicable

Important Notes:
- UNO uses 0-based indexing: VBA Cells(1, 1) = getCellByPosition(0, 0)
- VBA Range("A1") = getCellRangeByName("A1")
- Focus on observable behavior (cell values, not internal state)
- Tests should be independent and self-contained

Respond with JSON in this exact format:
{{
    "tests": [
        {{
            "test_name": "test_descriptive_name",
            "description": "What this test validates",
            "setup_code": "# Python code to set up test\\nsheet.getCellByPosition(0, 0).setValue(10)",
            "vba_expected_behavior": "Sets cell A1 to value 10, then multiplies by 2",
            "python_test_code": "# Call the translated function\\nTestFunction()",
            "assertions": [
                "assert sheet.getCellByPosition(0, 0).getValue() == 20",
                "assert sheet.getCellByPosition(0, 1).getString() == 'Done'"
            ]
        }}
    ]
}}

Generate {num_tests} comprehensive tests covering different aspects of the translation.
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
