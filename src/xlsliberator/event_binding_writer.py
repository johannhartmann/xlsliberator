"""Source-map-aware event binding rewrites for ODS content.xml."""

from __future__ import annotations

import re

from loguru import logger

from xlsliberator.validation_models import EventBindingIR


def rewrite_event_bindings(
    content_xml: str,
    py_modules: dict[str, str],
    event_bindings: list[EventBindingIR] | None = None,
) -> tuple[str, list[EventBindingIR]]:
    """Rewrite Basic event bindings to Python script URLs.

    Explicit ``event_bindings`` take precedence over heuristic module matching.
    Returned bindings contain unresolved entries when no rewrite was possible.
    """
    unresolved: list[EventBindingIR] = []
    updated = content_xml

    if event_bindings:
        for binding in event_bindings:
            if not binding.target_script_uri:
                unresolved.append(binding)
                continue
            if binding.source_handler and binding.source_handler in updated:
                updated = updated.replace(binding.source_handler, binding.target_script_uri)
            elif binding.source_handler:
                unresolved.append(binding)
        return updated, unresolved

    updated = _rewrite_vba_to_python_event_handlers(content_xml, py_modules)
    return updated, unresolved


def _rewrite_vba_to_python_event_handlers(content_xml: str, py_modules: dict[str, str]) -> str:
    """Heuristic compatibility rewrite for legacy Basic script URLs."""
    vba_to_py: dict[str, str] = {}
    for py_module_name in py_modules:
        if py_module_name.endswith(".py"):
            base_name = py_module_name[:-3]
            vba_to_py[base_name] = py_module_name
            if "." in base_name:
                module_only = base_name.split(".")[0]
                vba_to_py[module_only] = py_module_name

    if not vba_to_py:
        logger.debug("No VBA modules to map, skipping event handler rewriting")
        return content_xml

    pattern = (
        r"(vnd\.sun\.star\.script:)VBAProject\.([^?]+)\?language=Basic(&amp;location=document)"
    )

    def replace_handler(match: re.Match[str]) -> str:
        prefix = match.group(1)
        vba_path = match.group(2)
        location = match.group(3)
        original = match.group(0)
        parts = vba_path.split(".")
        if len(parts) < 2:
            logger.warning(f"Invalid VBA path format: {vba_path}")
            return original

        py_module = None
        function_name = ""
        for i in range(len(parts) - 1, 0, -1):
            vba_module = ".".join(parts[:i])
            if vba_module in vba_to_py:
                py_module = vba_to_py[vba_module]
                function_name = ".".join(parts[i:])
                break

        if not py_module:
            logger.warning(f"No Python module found for VBA path: {vba_path}")
            return original

        python_url = f"{prefix}{py_module}${function_name}?language=Python{location}"
        logger.debug(f"Rewrote event handler: {original} -> {python_url}")
        return python_url

    updated_content = re.sub(pattern, replace_handler, content_xml)
    replacements = len(re.findall(pattern, content_xml))
    if replacements > 0:
        logger.info(f"Rewrote {replacements} VBA event handlers to Python")
    return updated_content
