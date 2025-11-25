"""LLM-based VBA to Python-UNO translation using Claude API."""

import json
import os
import re
from pathlib import Path
from typing import TYPE_CHECKING

import yaml
from anthropic import Anthropic
from loguru import logger

from xlsliberator.python_syntax_validator import PythonSyntaxValidator

if TYPE_CHECKING:
    from xlsliberator.vba_reference_analyzer import VBAReferences
    from xlsliberator.vba_translation_validator import TranslationEvaluation


# Python-UNO Best Practices Context (from LibreOffice manual research)
PYTHON_UNO_BEST_PRACTICES = """
## Python-UNO Best Practices for LibreOffice Calc

### CRITICAL REQUIREMENTS

1. **g_exportedScripts Tuple**
   ALL macros MUST include a g_exportedScripts tuple listing callable functions:
   ```python
   def my_function():
       # ... code ...

   g_exportedScripts = (my_function,)  # Required!
   ```

2. **XSCRIPTCONTEXT for Document Access**
   Use XSCRIPTCONTEXT (provided by LibreOffice) to access the document:
   ```python
   def my_function():
       doc = XSCRIPTCONTEXT.getDocument()
       sheet = doc.Sheets.getByIndex(0)
       # ... work with sheet ...
   ```

### COMMON PATTERNS

**Access Cells:**
```python
cell = sheet.getCellRangeByName("A1")
cell.Value = 42              # number
cell.String = "text"         # string
cell.Formula = "=SUM(A1:A10)" # formula
```

**Get Active Sheet:**
```python
doc = XSCRIPTCONTEXT.getDocument()
sheet = doc.getCurrentController().getActiveSheet()
```

**Iterate Ranges:**
```python
cell_range = sheet.getCellRangeByName("A1:B10")
data = cell_range.getDataArray()  # tuple of tuples
for row in data:
    for cell_value in row:
        # ... process ...
```

**Call Calc Functions (FunctionAccess):**
```python
smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
fa = smgr.createInstanceWithContext(
    "com.sun.star.sheet.FunctionAccess",
    XSCRIPTCONTEXT.getComponentContext()
)
result = fa.callFunction("SUM", ((1, 2, 3),))  # tuple-of-sequences
```

**Event Handlers:**
Event handlers are just regular functions listed in g_exportedScripts.
Use function name to match the event (e.g., Workbook_Open).
```python
def Workbook_Open():
    doc = XSCRIPTCONTEXT.getDocument()
    # ... handle event ...

g_exportedScripts = (Workbook_Open,)
```

### STRICT PROHIBITIONS

**DO NOT USE:**
- subprocess, os.system, or any external process execution
- File I/O outside UNO APIs (no open(), file operations)
- External dependencies (requests, numpy, pandas, etc.)
- System-level operations
- Network operations outside UNO

**Instead use:**
- UNO APIs for all document operations
- com.sun.star.ucb.SimpleFileAccess for file operations
- FunctionAccess for Calc calculations
- Built-in Python stdlib only (no external packages)

### INDEX CONVERSION
VBA uses 1-based indexing, Python uses 0-based:
- VBA: Worksheets(1) → Python: doc.Sheets.getByIndex(0)
- VBA: Range cells → Python: adjust all indices by -1

### ERROR HANDLING
Replace VBA's "On Error Resume Next" with try-except:
```python
try:
    # ... risky operation ...
except Exception as e:
    logger.warning(f"Operation failed: {e}")
    # ... handle error ...
```

### LOGGING
Replace MsgBox/Debug.Print with logger:
```python
from loguru import logger

logger.info("Information message")
logger.warning("Warning message")
logger.error("Error message")
```
"""


class LLMVBATranslator:
    """Translates VBA code to Python-UNO using Claude LLM with rule-based mapping injection."""

    def __init__(
        self,
        vba_api_map_path: Path | None = None,
        event_map_path: Path | None = None,
        cache_path: Path | None = None,
    ):
        """Initialize LLM VBA translator.

        Args:
            vba_api_map_path: Path to VBA API mapping YAML
            event_map_path: Path to event mapping YAML
            cache_path: Optional path to cache translated VBA
        """
        self.client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        self.cache_path = cache_path or Path(".vba_cache.json")
        self.cache: dict[str, str] = self._load_cache()

        # Load mapping rules
        self.vba_api_map_path = vba_api_map_path or Path("rules/vba_api_map.yaml")
        self.event_map_path = event_map_path or Path("rules/event_map.yaml")
        self.vba_mappings = self._load_vba_mappings()
        self.event_mappings = self._load_event_mappings()

        # Initialize syntax validator
        self.syntax_validator = PythonSyntaxValidator()

    def _load_cache(self) -> dict[str, str]:
        """Load translation cache from disk."""
        if self.cache_path.exists():
            try:
                with open(self.cache_path) as f:
                    cache_data: dict[str, str] = json.load(f)
                    return cache_data
            except Exception as e:
                logger.warning(f"Failed to load VBA cache: {e}")
        return {}

    def _save_cache(self) -> None:
        """Save translation cache to disk."""
        try:
            with open(self.cache_path, "w") as f:
                json.dump(self.cache, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save VBA cache: {e}")

    def _load_vba_mappings(self) -> dict:
        """Load VBA API mapping rules from YAML."""
        if not self.vba_api_map_path.exists():
            logger.warning(f"VBA API mapping file not found: {self.vba_api_map_path}")
            return {}

        try:
            with open(self.vba_api_map_path) as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            logger.error(f"Failed to load VBA API mappings: {e}")
            return {}

    def _load_event_mappings(self) -> dict:
        """Load event mapping rules from YAML."""
        if not self.event_map_path.exists():
            logger.warning(f"Event mapping file not found: {self.event_map_path}")
            return {}

        try:
            with open(self.event_map_path) as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            logger.error(f"Failed to load event mappings: {e}")
            return {}

    def translate_vba(
        self, vba_code: str, is_event_handler: bool = False, enable_reference_aware: bool = True
    ) -> str:
        """Translate VBA code to Python-UNO using Claude LLM.

        Args:
            vba_code: VBA source code
            is_event_handler: Whether this is an event handler (affects context setup)
            enable_reference_aware: Use reference-aware prompts (Strategy A)

        Returns:
            Translated Python-UNO code
        """
        # Check cache first
        cache_key = f"{vba_code}:{is_event_handler}:{enable_reference_aware}"
        if cache_key in self.cache:
            logger.debug(f"LLM VBA cache hit for code: {vba_code[:50]}...")
            return self.cache[cache_key]

        # Call Claude API for translation
        logger.info(f"LLM VBA translation for code: {vba_code[:50]}...")

        # Analyze references if reference-aware enabled
        references = None
        if enable_reference_aware:
            from xlsliberator.vba_reference_analyzer import (
                analyze_vba_references,
            )

            references = analyze_vba_references(vba_code)
            logger.info(
                f"Reference-aware translation: {len(references.api_calls)} APIs, "
                f"{len(references.special_patterns)} patterns"
            )

        prompt = self._build_translation_prompt(vba_code, is_event_handler, references)

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=20000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}],
            )

            # Extract text from response
            content_block = response.content[0]
            if hasattr(content_block, "text"):
                translated: str = content_block.text.strip()
            else:
                raise ValueError(f"Unexpected content block type: {type(content_block)}")

            # Strip markdown code blocks if present
            translated = re.sub(r"^```python\n|^```\n|```$", "", translated, flags=re.MULTILINE)

            # Validate syntax (Phase 2: Syntax Validation)
            validation_result = self.syntax_validator.validate_syntax(translated)

            if not validation_result.is_valid:
                logger.error(
                    f"Syntax validation failed for translated code:\n"
                    f"Errors: {validation_result.syntax_errors}"
                )
                for error in validation_result.syntax_errors:
                    logger.error(f"  - {error}")

            if validation_result.warnings:
                logger.warning(f"Syntax validation warnings ({len(validation_result.warnings)}):")
                for warning in validation_result.warnings:
                    logger.warning(f"  - {warning}")

            if not validation_result.uno_compatible:
                logger.warning("Translated code may not be compatible with LibreOffice Python")

            # Cache the result
            self.cache[cache_key] = translated
            self._save_cache()

            logger.info(
                f"LLM VBA translation: {vba_code[:50]}... → {translated[:50]}... "
                f"(valid: {validation_result.is_valid}, "
                f"warnings: {len(validation_result.warnings)})"
            )
            return translated

        except Exception as e:
            logger.error(f"LLM VBA translation failed: {e}")
            # Fallback: return comment with original VBA
            return f"# VBA translation failed\n# Original VBA:\n# {vba_code}"

    def _build_translation_prompt(
        self, vba_code: str, _is_event_handler: bool, references: "VBAReferences | None" = None
    ) -> str:
        """Build prompt for Claude to translate VBA code.

        Args:
            vba_code: VBA code to translate
            _is_event_handler: Whether this is an event handler (reserved for future use)
            references: Optional VBA references for targeted translation

        Returns:
            Prompt string for Claude
        """
        # Extract relevant mapping sections
        object_mappings = self.vba_mappings.get("object_mappings", {})
        method_mappings = self.vba_mappings.get("method_mappings", {})
        function_mappings = self.vba_mappings.get("function_mappings", {})
        control_flow = self.vba_mappings.get("control_flow", {})
        declarations = self.vba_mappings.get("declarations", {})
        special_cases = self.vba_mappings.get("special_cases", {})
        required_imports = self.vba_mappings.get("required_imports", [])
        context_setup = self.vba_mappings.get("context_setup", "")

        # Filter mappings based on detected references (Strategy A)
        if references:
            object_mappings = self._filter_object_mappings(object_mappings, references)
            method_mappings = self._filter_method_mappings(method_mappings, references)
            control_flow = self._filter_control_flow(control_flow, references)

        # Format mappings for prompt
        object_map_str = self._format_mappings_for_prompt(object_mappings, "Object Mappings")
        method_map_str = self._format_method_mappings(method_mappings)
        function_map_str = self._format_mappings_for_prompt(function_mappings, "Function Mappings")
        control_flow_str = self._format_control_flow(control_flow)
        declarations_str = self._format_control_flow(declarations)
        special_str = self._format_control_flow(special_cases)

        imports_str = "\n".join(required_imports)

        # Build API usage section if references available
        api_usage_str = ""
        patterns_str = ""
        if references:
            api_usage_str = self._format_api_usage(references)
            patterns_str = self._format_special_patterns(references)

        # Build reference-aware prompt sections
        reference_sections = ""
        if references and (api_usage_str or patterns_str):
            reference_sections = f"""
{api_usage_str}

{patterns_str}

Translation Rules (filtered for detected APIs):
"""
        else:
            reference_sections = "Translation Rules:"

        prompt = f"""Translate this VBA code to Python-UNO format for LibreOffice Calc.

{PYTHON_UNO_BEST_PRACTICES}

VBA Code:
```vba
{vba_code}
```

{reference_sections}

{object_map_str}

{method_map_str}

{function_map_str}

{control_flow_str}

{declarations_str}

{special_str}

Required Imports:
```python
{imports_str}
```

Context Setup (include at start of translated code):
```python
{context_setup}
```

Requirements:
1. Follow the mapping rules above for all VBA constructs
2. Preserve the logic and behavior of the original VBA code
3. Use proper Python indentation (4 spaces)
4. Add type hints where appropriate
5. Handle errors with try-except blocks (not On Error)
6. Replace VBA comments (') with Python comments (#)
7. Convert VBA string literals to Python strings
8. Handle 1-based indexing → 0-based indexing for arrays/ranges
9. Add docstrings for functions/methods
10. Use logger.info() instead of MsgBox

Output ONLY the translated Python code, no explanations or markdown code blocks.

Example Translation:
VBA:
```vba
Sub UpdateCell()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Range("A1").Value = "Hello"
    MsgBox "Done"
End Sub
```

Python:
```python
import uno
from loguru import logger

def UpdateCell(*args):
    '''Auto-translated VBA procedure.'''
    # Get UNO context
    ctx = XSCRIPTCONTEXT.getComponentContext() if 'XSCRIPTCONTEXT' in dir() else uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = XSCRIPTCONTEXT.getDocument() if 'XSCRIPTCONTEXT' in dir() else desktop.getCurrentComponent()
    sheet = doc.getCurrentController().getActiveSheet()

    ws = None
    ws = sheet
    sheet.getCellRangeByName("A1").setString("Hello")
    logger.info("Done")
```

Now translate the VBA code above:"""

        return prompt

    def _format_mappings_for_prompt(self, mappings: dict, title: str) -> str:
        """Format simple key-value mappings for prompt."""
        if not mappings:
            return ""

        lines = [f"{title}:"]
        for vba_name, python_equiv in mappings.items():
            lines.append(f"  {vba_name} → {python_equiv}")

        return "\n".join(lines)

    def _format_method_mappings(self, mappings: dict) -> str:
        """Format method mappings with patterns and replacements."""
        if not mappings:
            return ""

        lines = ["Method/Property Mappings:"]
        for name, mapping in mappings.items():
            if isinstance(mapping, dict):
                pattern = mapping.get("pattern", "")
                replacement = mapping.get("replacement", "")
                description = mapping.get("description", "")
                lines.append(f"  {name}:")
                lines.append(f"    Pattern: {pattern}")
                lines.append(f"    Replacement: {replacement}")
                if description:
                    lines.append(f"    Description: {description}")

        return "\n".join(lines)

    def _format_control_flow(self, mappings: dict) -> str:
        """Format control flow mappings."""
        if not mappings:
            return ""

        lines = ["Control Flow / Special Cases:"]
        for name, mapping in mappings.items():
            if isinstance(mapping, dict):
                pattern = mapping.get("pattern", "")
                replacement = mapping.get("replacement", "")
                description = mapping.get("description", "")
                lines.append(f"  {name}: {pattern} → {replacement}")
                if description:
                    lines.append(f"    ({description})")

        return "\n".join(lines)

    def translate_event_handler(self, event_name: str, vba_code: str) -> str:
        """Translate VBA event handler to Python-UNO event handler.

        Args:
            event_name: VBA event name (e.g., "Workbook_Open")
            vba_code: VBA event handler code

        Returns:
            Python-UNO event handler code
        """
        # Check if this is a workbook or worksheet event
        workbook_events = self.event_mappings.get("workbook_events", {})
        worksheet_events = self.event_mappings.get("worksheet_events", {})

        event_info = None

        if event_name in workbook_events:
            event_info = workbook_events[event_name]
        elif event_name in worksheet_events:
            event_info = worksheet_events[event_name]

        if not event_info:
            logger.warning(f"Unknown event: {event_name}, using generic translation")
            return self.translate_vba(vba_code, is_event_handler=True)

        python_name = event_info.get("python_name", event_name.lower())
        description = event_info.get("description", "")

        # Translate the VBA body
        translated_body = self.translate_vba(vba_code, is_event_handler=True)

        # Wrap in event handler signature
        handler_code = f'''def {python_name}(event=None):
    """Event handler for {event_name}.

    {description}
    """
    import uno
    from loguru import logger

{self._indent_code(translated_body, 1)}
'''

        return handler_code

    def _indent_code(self, code: str, levels: int) -> str:
        """Indent code by specified levels.

        Args:
            code: Code to indent
            levels: Number of indent levels (4 spaces each)

        Returns:
            Indented code
        """
        indent = "    " * levels
        lines = code.split("\n")
        return "\n".join(indent + line if line.strip() else "" for line in lines)

    def _filter_object_mappings(self, mappings: dict, references: "VBAReferences") -> dict:
        """Filter object mappings to only those referenced in VBA code.

        Args:
            mappings: Full object mappings dict
            references: Detected VBA references

        Returns:
            Filtered mappings dict
        """

        filtered: dict[str, str] = {}

        for api_name in references.api_calls:
            if api_name in mappings:
                filtered[api_name] = mappings[api_name]

        # Always include commonly related APIs
        related_apis = {
            "Range": ["Cells", "ActiveSheet", "Worksheets"],
            "Cells": ["Range", "ActiveSheet"],
            "Worksheets": ["ActiveSheet", "Workbooks"],
        }

        for detected_api in references.api_calls:
            if detected_api in related_apis:
                for related in related_apis[detected_api]:
                    if related in mappings:
                        filtered[related] = mappings[related]

        return filtered if filtered else mappings

    def _filter_method_mappings(self, mappings: dict, references: "VBAReferences") -> dict:
        """Filter method mappings to only those referenced in VBA code.

        Args:
            mappings: Full method mappings dict
            references: Detected VBA references

        Returns:
            Filtered mappings dict
        """

        filtered: dict[str, dict] = {}

        for method_name, method_config in mappings.items():
            # Check if any detected API is related to this method
            for api_name in references.api_calls:
                if api_name.lower() in method_name.lower():
                    filtered[method_name] = method_config
                    break

        return filtered if filtered else mappings

    def _filter_control_flow(self, mappings: dict, references: "VBAReferences") -> dict:
        """Filter control flow mappings based on detected patterns.

        Args:
            mappings: Full control flow mappings dict
            references: Detected VBA references

        Returns:
            Filtered mappings dict
        """

        filtered: dict[str, dict] = {}

        pattern_to_mapping = {
            "for_each_loop": ["For_Each"],
            "for_to_loop": ["For_To", "Next"],
            "do_while_loop": ["Do_While", "Loop"],
            "select_case": ["Select_Case"],
            "with_block": ["With"],
            "exit_early": ["Exit_Sub", "Exit_Function"],
        }

        for pattern in references.special_patterns:
            if pattern in pattern_to_mapping:
                for mapping_key in pattern_to_mapping[pattern]:
                    if mapping_key in mappings:
                        filtered[mapping_key] = mappings[mapping_key]

        return filtered if filtered else mappings

    def _format_api_usage(self, references: "VBAReferences") -> str:
        """Format API usage statistics for prompt.

        Args:
            references: Detected VBA references

        Returns:
            Formatted API usage string
        """

        if not references.api_calls:
            return ""

        sorted_apis = sorted(references.api_calls.items(), key=lambda x: x[1], reverse=True)

        lines = ["Detected API Usage (prioritize these mappings):"]
        for api, count in sorted_apis[:10]:
            lines.append(f"  {api}: {count} occurrence{'s' if count > 1 else ''}")

        return "\n".join(lines)

    def _format_special_patterns(self, references: "VBAReferences") -> str:
        """Format special patterns for prompt.

        Args:
            references: Detected VBA references

        Returns:
            Formatted special patterns string
        """

        if not references.special_patterns:
            return ""

        pattern_descriptions = {
            "error_handling": "Error handling (On Error Resume Next/GoTo)",
            "for_each_loop": "For Each loops",
            "for_to_loop": "For...To loops",
            "do_while_loop": "Do While/Until loops",
            "arrays": "Array declarations",
            "dynamic_arrays": "Dynamic arrays (ReDim)",
            "select_case": "Select Case statements",
            "with_block": "With blocks",
            "exit_early": "Early exit (Exit Sub/Function)",
            "string_concatenation": "Complex string concatenation",
            "optional_params": "Optional parameters",
            "byref_params": "ByRef parameters",
            "property_procedures": "Property Get/Let/Set",
            "late_binding": "Late binding (CreateObject/GetObject)",
            "worksheet_functions": "WorksheetFunction calls",
            "variant_types": "Variant types",
            "object_variables": "Object variables (Set keyword)",
            "user_defined_types": "User-defined types",
            "named_parameters": "Named parameters (:=)",
        }

        lines = ["Special Patterns Detected:"]
        for pattern in references.special_patterns:
            description = pattern_descriptions.get(pattern, pattern)
            lines.append(f"  - {description}")

        return "\n".join(lines)

    def translate_vba_with_reflection(
        self,
        vba_code: str,
        is_event_handler: bool = False,
        max_iterations: int = 3,
        quality_threshold: int = 85,
    ) -> tuple[str, list]:
        """Translate VBA with agentic reflection loop (Phase 3).

        Iteratively improves translation quality by:
        1. Initial translation (reference-aware)
        2. Self-evaluation by LLM
        3. Refinement based on feedback
        4. Repeat until quality threshold met or max iterations

        Args:
            vba_code: VBA source code
            is_event_handler: Whether this is an event handler
            max_iterations: Maximum refinement iterations (default: 3)
            quality_threshold: Minimum acceptable quality 0-100 (default: 85)

        Returns:
            Tuple of (final_python_code, evaluation_history)
        """
        from xlsliberator.vba_reference_analyzer import analyze_vba_references
        from xlsliberator.vba_translation_validator import VBATranslationValidator

        logger.info(f"Starting reflection-based translation (max {max_iterations} iterations)")

        references = analyze_vba_references(vba_code)
        validator = VBATranslationValidator()

        evaluations = []

        # Initial translation (reference-aware, Phase 1)
        python_code = self.translate_vba(vba_code, is_event_handler, enable_reference_aware=True)

        for iteration in range(max_iterations):
            logger.info(f"Reflection iteration {iteration + 1}/{max_iterations}")

            # Evaluate translation (Phase 3: Self-evaluation)
            evaluation = validator.evaluate_translation(vba_code, python_code, references)
            evaluations.append(evaluation)

            logger.info(
                f"Translation quality: {evaluation.overall_quality}/100 "
                f"(acceptable: {evaluation.is_acceptable}, issues: {len(evaluation.issues)})"
            )

            # Log issues
            if evaluation.issues:
                logger.warning(f"Issues found ({len(evaluation.issues)}):")
                for issue in evaluation.issues:
                    logger.warning(f"  - [{issue.issue_type.value}] {issue.description}")

            # Check if acceptable
            if evaluation.is_acceptable and evaluation.overall_quality >= quality_threshold:
                logger.success(
                    f"Translation accepted (quality: {evaluation.overall_quality}/100, "
                    f"iteration: {iteration + 1})"
                )
                break

            # If not acceptable and not last iteration, refine
            if iteration < max_iterations - 1:
                logger.info(f"Refining translation (quality below threshold: {quality_threshold})")
                python_code = self._refine_translation(
                    vba_code, python_code, evaluation, references
                )
            else:
                logger.warning(
                    f"Max iterations reached, returning best translation "
                    f"(quality: {evaluation.overall_quality}/100)"
                )

        return python_code, evaluations

    def _refine_translation(
        self,
        vba_code: str,
        python_code: str,
        evaluation: "TranslationEvaluation",
        references: "VBAReferences",
    ) -> str:
        """Refine translation based on evaluation feedback.

        Args:
            vba_code: Original VBA code
            python_code: Current Python translation
            evaluation: TranslationEvaluation with issues and suggestions
            references: VBA references

        Returns:
            Improved Python-UNO code
        """
        prompt = self._build_refinement_prompt(vba_code, python_code, evaluation, references)

        try:
            response = self.client.messages.create(
                model="claude-sonnet-4-5",
                max_tokens=20000,
                temperature=0,
                messages=[{"role": "user", "content": prompt}],
            )

            content_block = response.content[0]
            if hasattr(content_block, "text"):
                refined_code: str = content_block.text.strip()
            else:
                raise ValueError(f"Unexpected content block type: {type(content_block)}")

            # Strip markdown code blocks if present
            refined_code = re.sub(r"^```python\n|^```\n|```$", "", refined_code, flags=re.MULTILINE)

            logger.info(f"Translation refined: {len(python_code)} → {len(refined_code)} bytes")
            return refined_code

        except Exception as e:
            logger.error(f"Translation refinement failed: {e}")
            # Return original if refinement fails
            return python_code

    def _build_refinement_prompt(
        self,
        vba_code: str,
        python_code: str,
        evaluation: "TranslationEvaluation",
        references: "VBAReferences",
    ) -> str:
        """Build prompt for translation refinement.

        Args:
            vba_code: Original VBA code
            python_code: Current Python translation
            evaluation: TranslationEvaluation with feedback
            references: VBA references

        Returns:
            Refinement prompt string
        """
        issues_str = self._format_issues_for_prompt(evaluation.issues)
        suggestions_str = "\n".join(f"- {s}" for s in evaluation.suggestions)
        api_usage_str = self._format_api_usage(references)

        return f"""Improve this VBA-to-Python-UNO translation based on evaluation feedback.

Original VBA Code:
```vba
{vba_code}
```

Current Python Translation (Quality: {evaluation.overall_quality}/100):
```python
{python_code}
```

Detected VBA API Usage:
{api_usage_str}

Issues to Fix ({len(evaluation.issues)}):
{issues_str}

Improvement Suggestions:
{suggestions_str}

Your task:
1. Fix all identified issues
2. Ensure all VBA logic is preserved
3. Use correct Python-UNO API mappings
4. Fix indexing errors (VBA 1-based → Python 0-based)
5. Add proper error handling
6. Include all required imports

Provide an improved Python-UNO translation.
Output ONLY the corrected Python code, no explanations or markdown.
"""

    def _format_issues_for_prompt(self, issues: list) -> str:
        """Format issues for refinement prompt.

        Args:
            issues: List of TranslationIssue objects

        Returns:
            Formatted string
        """
        if not issues:
            return "(No issues)"

        lines = []
        for i, issue in enumerate(issues, 1):
            lines.append(f"{i}. [{issue.issue_type.value}] (severity {issue.severity}/10)")
            lines.append(f"   Description: {issue.description}")
            if issue.vba_line:
                lines.append(f"   VBA line: {issue.vba_line}")
            if issue.python_line:
                lines.append(f"   Python line: {issue.python_line}")

        return "\n".join(lines)
