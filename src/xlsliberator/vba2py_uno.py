"""VBA to Python-UNO translator (Phase F8 - Minimal Subset)."""

import os  # noqa: F401
import re
from dataclasses import dataclass

from loguru import logger  # noqa: F401


class VBATranslationError(Exception):
    """Raised when VBA translation fails."""


@dataclass
class TranslationResult:
    """Result of VBA to Python translation."""

    python_code: str
    warnings: list[str]
    unsupported_features: list[str]


def translate_vba_to_python(vba_code: str, use_llm: bool = True) -> TranslationResult:
    """Translate VBA code to Python-UNO code (hybrid approach).

    Args:
        vba_code: VBA source code
        use_llm: Use LLM-based translation as primary (True) or rule-based only (False)

    Returns:
        TranslationResult with Python code and warnings

    Note:
        Hybrid approach:
        - If use_llm=True and ANTHROPIC_API_KEY set: Use LLM translation with mapping injection
        - Otherwise: Use rule-based translation (Phase F8 minimal)

        Rule-based supports:
        - Sub/Function declarations
        - Dim statements
        - Range/Cells/Worksheets API calls
        - Simple assignments
        - MsgBox → logger
        - Comments

    Not supported in rule-based (will warn):
        - Complex control flow (For/If/Select beyond basic)
        - Error handling (On Error)
        - Arrays, UDTs
        - COM objects
    """
    # Try LLM translation if enabled
    if use_llm and os.environ.get("ANTHROPIC_API_KEY"):
        logger.info("Using LLM-based VBA translation (Claude with mapping injection)")
        try:
            from xlsliberator.llm_vba_translator import LLMVBATranslator  # noqa: F401

            translator = LLMVBATranslator()
            python_code = translator.translate_vba(vba_code, is_event_handler=False)
            return TranslationResult(python_code=python_code, warnings=[], unsupported_features=[])
        except Exception as e:
            logger.warning(f"LLM VBA translation failed, falling back to rule-based: {e}")
            # Fall through to rule-based translation

    # Rule-based translation (fallback or primary if use_llm=False)
    logger.info("Using rule-based VBA translation")
    warnings: list[str] = []
    unsupported: list[str] = []

    lines = vba_code.split("\n")
    python_lines = []
    indent_level = 0

    # Add standard imports
    python_lines.append("# Auto-translated from VBA")
    python_lines.append("import uno")
    python_lines.append("from loguru import logger")
    python_lines.append("")

    for line in lines:
        stripped = line.strip()

        # Skip empty lines and attributes
        if not stripped or stripped.startswith("Attribute "):
            continue

        # Comments
        if stripped.startswith("'") or stripped.startswith("Rem "):
            python_lines.append("    " * indent_level + "# " + stripped.lstrip("'Rem "))
            continue

        # Sub/Function declarations
        if re.match(r"(?:Public |Private )?(?:Sub|Function)\s+\w+", stripped, re.IGNORECASE):
            func_match = re.search(r"(Sub|Function)\s+(\w+)", stripped, re.IGNORECASE)
            if func_match:
                func_name = func_match.group(2)
                python_lines.append(f"def {func_name}(*args):")
                indent_level = 1
                python_lines.append("    " + '"""Auto-translated VBA procedure."""')
            continue

        # End Sub/Function
        if re.match(r"End\s+(Sub|Function)", stripped, re.IGNORECASE):
            indent_level = 0
            python_lines.append("")
            continue

        # Dim statements
        if stripped.lower().startswith("dim "):
            # Simple conversion: Dim x As Type → x = None
            var_match = re.search(r"Dim\s+(\w+)", stripped, re.IGNORECASE)
            if var_match:
                var_name = var_match.group(1)
                python_lines.append("    " * indent_level + f"{var_name} = None")
            continue

        # Range/Cells/Worksheets translations
        translated_line = _translate_excel_api(stripped)

        # MsgBox → logger
        translated_line = re.sub(
            r'MsgBox\s+"([^"]*)"', r'logger.info("\1")', translated_line, flags=re.IGNORECASE
        )
        translated_line = re.sub(
            r"MsgBox\s+(\w+)", r"logger.info(str(\1))", translated_line, flags=re.IGNORECASE
        )

        # DoEvents → pass
        if "DoEvents" in translated_line:
            translated_line = re.sub(r"\bDoEvents\b", "pass", translated_line, flags=re.IGNORECASE)

        if translated_line and translated_line != stripped:
            python_lines.append("    " * max(1, indent_level) + translated_line)
        elif stripped and not any(kw in stripped.lower() for kw in ["dim ", "attribute", "option"]):
            # Unsupported statement
            python_lines.append("    " * max(1, indent_level) + f"# TODO: {stripped}")
            unsupported.append(f"Unsupported: {stripped}")

    python_code = "\n".join(python_lines)

    return TranslationResult(
        python_code=python_code, warnings=warnings, unsupported_features=unsupported
    )


def _translate_excel_api(vba_line: str) -> str:
    """Translate Excel API calls from VBA to Python-UNO.

    Args:
        vba_line: Single line of VBA code

    Returns:
        Translated Python code
    """
    line = vba_line

    # Get UNO context (add if needed)
    if any(
        api in line for api in ["Range", "Cells", "Worksheets", "ActiveSheet", "ActiveWorkbook"]
    ):
        # These need UNO document context
        # For simplicity, assume doc is available
        pass

    # Range("A1") → sheet.getCellRangeByName("A1")
    line = re.sub(
        r'Range\("([A-Z0-9:]+)"\)', r'sheet.getCellRangeByName("\1")', line, flags=re.IGNORECASE
    )

    # Cells(row, col) → sheet.getCellByPosition(col-1, row-1)
    line = re.sub(
        r"Cells\((\d+),\s*(\d+)\)",
        lambda m: f"sheet.getCellByPosition({int(m.group(2)) - 1}, {int(m.group(1)) - 1})",
        line,
        flags=re.IGNORECASE,
    )

    # Worksheets("Name") → doc.getSheets().getByName("Name")
    line = re.sub(
        r'Worksheets\("([^"]+)"\)', r'doc.getSheets().getByName("\1")', line, flags=re.IGNORECASE
    )

    # ActiveSheet → doc.getCurrentController().getActiveSheet()
    line = re.sub(
        r"\bActiveSheet\b", "doc.getCurrentController().getActiveSheet()", line, flags=re.IGNORECASE
    )

    # .Value = → .setValue() or .setString()
    line = re.sub(r'\.Value\s*=\s*"([^"]*)"', r'.setString("\1")', line)
    line = re.sub(r"\.Value\s*=\s*(\d+(?:\.\d+)?)", r".setValue(\1)", line)

    return line


def create_event_handler_stub(event_name: str, vba_code: str, use_llm: bool = True) -> str:
    """Create Python-UNO event handler from VBA code.

    Args:
        event_name: Event name (e.g., "Workbook_Open")
        vba_code: VBA event handler code
        use_llm: Use LLM-based translation (hybrid approach)

    Returns:
        Python-UNO event handler code

    Note:
        Phase F8 - creates minimal working event handlers.
        Maps Workbook_Open → on_open, etc.

        Hybrid approach:
        - If use_llm=True and ANTHROPIC_API_KEY set: Use LLM event handler translator
        - Otherwise: Use rule-based translation
    """
    # Try LLM event handler translation if enabled
    if use_llm and os.environ.get("ANTHROPIC_API_KEY"):
        try:
            from xlsliberator.llm_vba_translator import LLMVBATranslator  # noqa: F401

            translator = LLMVBATranslator()
            return translator.translate_event_handler(event_name, vba_code)
        except Exception as e:
            logger.warning(f"LLM event handler translation failed, falling back to rule-based: {e}")
            # Fall through to rule-based translation

    # Rule-based translation (fallback or primary if use_llm=False)
    # Map VBA event names to Python UNO conventions
    event_map = {
        "Workbook_Open": "on_open",
        "Workbook_BeforeClose": "on_before_close",
        "Worksheet_Change": "on_worksheet_change",
    }

    python_name = event_map.get(event_name, event_name.lower())

    # Translate the VBA code
    result = translate_vba_to_python(vba_code, use_llm=False)

    # Wrap in proper event handler signature
    handler_code = f"""def {python_name}(event=None):
    '''Event handler for {event_name}.'''
    import uno

    # Get document context
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    # Get active sheet
    sheet = doc.getCurrentController().getActiveSheet()

    # Translated VBA code
{_indent_code(result.python_code, 1)}
"""

    return handler_code


def _indent_code(code: str, levels: int) -> str:
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
