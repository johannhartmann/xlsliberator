"""VBA reference analysis for targeted translation.

Analyzes VBA code to detect API usage, dependencies, and special patterns
to enable reference-aware LLM translation prompts.

Uses hybrid approach:
- oletools (existing) for simple extraction: API calls, procedures, dependencies
- LLM (Claude) for complex pattern detection: loops, error handling, etc.
- Regex fallback if LLM unavailable
"""

import json
import os
import re
from dataclasses import dataclass, field

from loguru import logger

from xlsliberator.extract_vba import (
    _extract_api_calls,
    _extract_dependencies,
    _extract_procedures,
)


@dataclass
class VBAReferences:
    """Detected VBA references for targeted translation."""

    api_calls: dict[str, int] = field(default_factory=dict)  # API -> usage count
    dependencies: set[str] = field(default_factory=set)  # Referenced modules
    procedures: list[str] = field(default_factory=list)  # Called procedures
    special_patterns: list[str] = field(default_factory=list)  # E.g., "error_handling"


def analyze_vba_references(vba_code: str) -> VBAReferences:
    """Analyze VBA code for all references and patterns.

    Args:
        vba_code: VBA source code to analyze

    Returns:
        VBAReferences with detected APIs, dependencies, and patterns

    Example:
        >>> refs = analyze_vba_references('''
        ... Sub Test()
        ...     On Error Resume Next
        ...     Dim arr(10) As Integer
        ...     For Each cell In Range("A1:A10")
        ...         cell.Value = 100
        ...     Next cell
        ... End Sub
        ... ''')
        >>> "error_handling" in refs.special_patterns
        True
        >>> "arrays" in refs.special_patterns
        True
        >>> "for_each_loop" in refs.special_patterns
        True
    """
    logger.debug(f"Analyzing VBA references for code: {vba_code[:50]}...")

    # Extract API calls using existing functionality
    api_calls = _extract_api_calls(vba_code)

    # Extract dependencies
    dependencies = _extract_dependencies(vba_code)

    # Extract procedures
    procedures = _extract_procedures(vba_code)

    # Detect special patterns
    special_patterns = _detect_special_patterns(vba_code)

    references = VBAReferences(
        api_calls=api_calls,
        dependencies=dependencies,
        procedures=procedures,
        special_patterns=special_patterns,
    )

    logger.info(
        f"VBA analysis: {len(api_calls)} APIs, {len(dependencies)} deps, "
        f"{len(procedures)} procs, {len(special_patterns)} patterns"
    )

    return references


def _llm_detect_patterns(vba_code: str) -> list[str]:
    """Detect special VBA patterns using Claude LLM.

    Args:
        vba_code: VBA source code

    Returns:
        List of detected pattern identifiers

    Raises:
        Exception: If LLM call fails
    """
    from anthropic import Anthropic

    client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    prompt = f"""Analyze this VBA code and detect special patterns.

VBA Code:
```vba
{vba_code}
```

Patterns to detect:
- error_handling: On Error Resume Next / On Error GoTo statements
- for_each_loop: For Each ... In loops
- for_to_loop: For i = 1 To n loops
- do_while_loop: Do While / Do Until loops
- arrays: Dim arr() As Type or array declarations
- dynamic_arrays: ReDim or ReDim Preserve
- select_case: Select Case statements
- with_block: With ... End With blocks
- exit_early: Exit Sub / Exit Function statements
- string_concatenation: Complex string building with multiple & operators
- optional_params: Optional keyword in parameters
- byref_params: ByRef keyword in parameters
- property_procedures: Property Get/Let/Set
- late_binding: CreateObject / GetObject calls
- worksheet_functions: WorksheetFunction.xxx calls
- variant_types: As Variant type declarations
- object_variables: Set keyword for object assignment
- user_defined_types: Type ... End Type declarations
- named_parameters: Named parameters using :=

Respond with JSON containing only the patterns you detect:
{{"patterns": ["for_each_loop", "error_handling", "arrays"]}}

IMPORTANT: Output ONLY the JSON object, no explanations or markdown."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=500,
            temperature=0,
            messages=[{"role": "user", "content": prompt}],
        )

        # Extract text from response
        content_block = response.content[0]
        if hasattr(content_block, "text"):
            json_text = content_block.text.strip()
        else:
            raise ValueError(f"Unexpected content block type: {type(content_block)}")

        # Strip markdown code blocks if present
        json_text = re.sub(r"^```json\n|^```\n|```$", "", json_text, flags=re.MULTILINE)

        # Parse JSON
        result: dict = json.loads(json_text)
        patterns: list[str] = result.get("patterns", [])

        logger.debug(f"LLM detected {len(patterns)} patterns: {patterns}")
        return patterns

    except Exception as e:
        logger.error(f"LLM pattern detection failed: {e}")
        raise


def _regex_detect_patterns(vba_code: str) -> list[str]:
    """Detect special VBA patterns using regex (fallback).

    This is a fallback when LLM is unavailable. Less reliable than LLM
    but better than nothing.

    Args:
        vba_code: VBA source code

    Returns:
        List of detected pattern identifiers
    """
    return _detect_special_patterns_regex(vba_code)


def _detect_special_patterns(vba_code: str) -> list[str]:
    """Detect special VBA patterns using hybrid approach.

    Tries LLM first for accuracy, falls back to regex if unavailable.

    Args:
        vba_code: VBA source code

    Returns:
        List of detected pattern identifiers
    """
    # Try LLM first (if API key available)
    if os.environ.get("ANTHROPIC_API_KEY"):
        try:
            return _llm_detect_patterns(vba_code)
        except Exception as e:
            logger.warning(
                f"LLM pattern detection failed: {e}, falling back to regex-based detection"
            )

    # Fallback to regex-based detection
    logger.info("Using regex-based pattern detection (LLM unavailable)")
    return _regex_detect_patterns(vba_code)


def _detect_special_patterns_regex(vba_code: str) -> list[str]:
    """Detect special VBA patterns that need careful translation.

    Args:
        vba_code: VBA source code

    Returns:
        List of detected pattern identifiers

    Detected Patterns:
        - error_handling: On Error Resume Next / On Error GoTo
        - for_each_loop: For Each ... In
        - for_to_loop: For i = 1 To n
        - do_while_loop: Do While / Do Until
        - arrays: Dim arr() As Type
        - select_case: Select Case
        - with_block: With ... End With
        - exit_early: Exit Sub / Exit Function
        - string_concatenation: Multiple & operators
        - optional_params: Optional keyword in params
        - byref_params: ByRef keyword in params
        - property_procedures: Property Get/Let/Set
        - late_binding: CreateObject / GetObject
        - worksheet_functions: WorksheetFunction.
        - variant_types: As Variant
    """
    patterns = []

    # Error handling
    if re.search(r"\bOn\s+Error\s+(Resume\s+Next|GoTo\s+\w+)", vba_code, re.IGNORECASE):
        patterns.append("error_handling")

    # Loop patterns
    if re.search(r"\bFor\s+Each\b", vba_code, re.IGNORECASE):
        patterns.append("for_each_loop")

    if re.search(r"\bFor\s+\w+\s*=\s*.+\s+To\s+", vba_code, re.IGNORECASE):
        patterns.append("for_to_loop")

    if re.search(r"\bDo\s+(While|Until)\b", vba_code, re.IGNORECASE):
        patterns.append("do_while_loop")

    # Arrays
    if re.search(r"\bDim\s+\w+\s*\(", vba_code, re.IGNORECASE):
        patterns.append("arrays")

    if re.search(r"\bReDim\s+(Preserve\s+)?\w+", vba_code, re.IGNORECASE):
        patterns.append("dynamic_arrays")

    # Select Case
    if re.search(r"\bSelect\s+Case\b", vba_code, re.IGNORECASE):
        patterns.append("select_case")

    # With blocks
    if re.search(r"\bWith\s+.+", vba_code, re.IGNORECASE):
        patterns.append("with_block")

    # Early exit
    if re.search(r"\bExit\s+(Sub|Function)\b", vba_code, re.IGNORECASE):
        patterns.append("exit_early")

    # String concatenation (3+ & operators suggests complex string building)
    ampersand_count = len(re.findall(r"\s&\s", vba_code))
    if ampersand_count >= 3:
        patterns.append("string_concatenation")

    # Optional parameters
    if re.search(r"\bOptional\s+\w+", vba_code, re.IGNORECASE):
        patterns.append("optional_params")

    # ByRef parameters
    if re.search(r"\bByRef\s+\w+", vba_code, re.IGNORECASE):
        patterns.append("byref_params")

    # Property procedures
    if re.search(r"\bProperty\s+(Get|Let|Set)\b", vba_code, re.IGNORECASE):
        patterns.append("property_procedures")

    # Late binding
    if re.search(r"\b(CreateObject|GetObject)\s*\(", vba_code, re.IGNORECASE):
        patterns.append("late_binding")

    # WorksheetFunction
    if re.search(r"\bWorksheetFunction\.", vba_code, re.IGNORECASE):
        patterns.append("worksheet_functions")

    # Variant types (often need special handling)
    if re.search(r"\bAs\s+Variant\b", vba_code, re.IGNORECASE):
        patterns.append("variant_types")

    # Object variables (Set keyword)
    if re.search(r"\bSet\s+\w+\s*=", vba_code, re.IGNORECASE):
        patterns.append("object_variables")

    # User-defined types
    if re.search(r"\bType\s+\w+", vba_code, re.IGNORECASE):
        patterns.append("user_defined_types")

    # Named parameters
    if re.search(r":=", vba_code):
        patterns.append("named_parameters")

    logger.debug(f"Detected special patterns: {patterns}")

    return patterns


def get_top_apis(references: VBAReferences, top_n: int = 10) -> list[tuple[str, int]]:
    """Get top N most-used APIs from references.

    Args:
        references: VBA references
        top_n: Number of top APIs to return

    Returns:
        List of (api_name, count) tuples, sorted by count descending
    """
    sorted_apis = sorted(references.api_calls.items(), key=lambda x: x[1], reverse=True)
    return sorted_apis[:top_n]


def get_translation_complexity_score(references: VBAReferences) -> int:
    """Calculate translation complexity score (0-100).

    Args:
        references: VBA references

    Returns:
        Complexity score where:
        - 0-30: Simple (basic API calls, no special patterns)
        - 31-60: Moderate (multiple APIs, some patterns)
        - 61-100: Complex (many APIs, multiple patterns)
    """
    score = 0

    # API diversity (max 30 points)
    api_count = len(references.api_calls)
    score += min(api_count * 3, 30)

    # Special patterns (max 40 points)
    pattern_count = len(references.special_patterns)
    score += min(pattern_count * 5, 40)

    # Dependencies (max 15 points)
    dep_count = len(references.dependencies)
    score += min(dep_count * 5, 15)

    # Procedures (max 15 points)
    proc_count = len(references.procedures)
    score += min(proc_count * 3, 15)

    return min(score, 100)


def get_recommended_translation_strategy(references: VBAReferences) -> str:
    """Recommend translation strategy based on complexity.

    Args:
        references: VBA references

    Returns:
        Strategy name: "rule_based", "llm_basic", or "llm_reflection"
    """
    complexity = get_translation_complexity_score(references)

    if complexity <= 30:
        return "rule_based"
    elif complexity <= 60:
        return "llm_basic"
    else:
        return "llm_reflection"
