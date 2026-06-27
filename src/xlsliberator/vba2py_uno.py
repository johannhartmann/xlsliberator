"""VBA to Python-UNO translator (LLM-based)."""

import os
import re
from dataclasses import dataclass

from loguru import logger


class VBATranslationError(Exception):
    """Raised when VBA translation fails."""


@dataclass
class TranslationResult:
    """Result of VBA to Python translation."""

    python_code: str
    warnings: list[str]
    unsupported_features: list[str]


def translate_vba_to_python(
    vba_code: str,
    module_name: str | None = None,
) -> TranslationResult:
    """Translate VBA code to Python-UNO code using the LLM translator.

    Args:
        vba_code: VBA source code
        module_name: Optional VBA module name for source-map markers

    Returns:
        TranslationResult with Python code and warnings

    Raises:
        VBATranslationError: If ANTHROPIC_API_KEY is not set.

    Note:
        LLM translation is the only supported path. Callers must ensure an
        ANTHROPIC_API_KEY is configured; ``api.convert`` skips VBA translation
        entirely when it is absent rather than embedding a degraded fallback.
    """
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise VBATranslationError("VBA translation requires ANTHROPIC_API_KEY to be set")

    logger.info("Using LLM-based VBA translation (Claude with mapping injection)")
    from xlsliberator.llm_vba_translator import LLMVBATranslator

    translator = LLMVBATranslator()
    python_code = translator.translate_vba(vba_code, is_event_handler=False)
    python_code = _inject_source_markers(
        python_code,
        _extract_vba_procedure_names(vba_code),
        module_name,
    )
    return TranslationResult(python_code=python_code, warnings=[], unsupported_features=[])


def _extract_vba_procedure_names(vba_code: str) -> list[str]:
    # Anchor to the start of a line and use horizontal whitespace only, so
    # "End Sub"/"Exit Sub" (which have a keyword before Sub) and the name on the
    # following line are not captured as spurious procedure names.
    return [
        match.group(1)
        for match in re.finditer(
            r"^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*(?:Sub|Function)[ \t]+(\w+)",
            vba_code,
            re.IGNORECASE | re.MULTILINE,
        )
    ]


def _source_marker(module_name: str | None, procedure: str) -> str:
    module = module_name or "unknown"
    artifact_id = f"{module}.{procedure}"
    return (
        f"# xlsliberator-source: module={module}; procedure={procedure}; artifact_id={artifact_id}"
    )


def _inject_source_markers(
    python_code: str,
    procedures: list[str],
    module_name: str | None,
) -> str:
    """Inject source-map markers after matching generated function definitions."""
    updated = python_code
    for procedure in procedures:
        # Match the full signature line including nested parens in defaults
        # (e.g. ``def foo(x=(1, 2)):``) and an optional return annotation. ``.``
        # stays within the line, so the greedy ``\)`` binds to the closing paren
        # before the colon.
        pattern = rf"(def\s+{re.escape(procedure)}\s*\(.*\)\s*(?:->[^\n:]+)?:[ \t]*\n)"
        replacement = rf"\1    {_source_marker(module_name, procedure)}\n"
        updated = re.sub(pattern, replacement, updated, count=1)
    return updated


def create_event_handler_stub(event_name: str, vba_code: str) -> str:
    """Create a Python-UNO event handler from VBA code using the LLM translator.

    Args:
        event_name: Event name (e.g., "Workbook_Open")
        vba_code: VBA event handler code

    Returns:
        Python-UNO event handler code

    Raises:
        VBATranslationError: If ANTHROPIC_API_KEY is not set.
    """
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise VBATranslationError("Event handler translation requires ANTHROPIC_API_KEY to be set")

    from xlsliberator.llm_vba_translator import LLMVBATranslator

    translator = LLMVBATranslator()
    return translator.translate_event_handler(event_name, vba_code)
