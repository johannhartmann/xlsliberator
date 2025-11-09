"""LLM-based VBA to Python-UNO translation using Claude API."""

import json
import os
from pathlib import Path

import yaml
from anthropic import Anthropic
from loguru import logger


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

    def translate_vba(self, vba_code: str, is_event_handler: bool = False) -> str:
        """Translate VBA code to Python-UNO using Claude LLM.

        Args:
            vba_code: VBA source code
            is_event_handler: Whether this is an event handler (affects context setup)

        Returns:
            Translated Python-UNO code
        """
        # Check cache first
        cache_key = f"{vba_code}:{is_event_handler}"
        if cache_key in self.cache:
            logger.debug(f"LLM VBA cache hit for code: {vba_code[:50]}...")
            return self.cache[cache_key]

        # Call Claude API for translation
        logger.info(f"LLM VBA translation for code: {vba_code[:50]}...")

        prompt = self._build_translation_prompt(vba_code, is_event_handler)

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

            # Cache the result
            self.cache[cache_key] = translated
            self._save_cache()

            logger.info(f"LLM VBA translation: {vba_code[:50]}... → {translated[:50]}...")
            return translated

        except Exception as e:
            logger.error(f"LLM VBA translation failed: {e}")
            # Fallback: return comment with original VBA
            return f"# VBA translation failed\n# Original VBA:\n# {vba_code}"

    def _build_translation_prompt(self, vba_code: str, _is_event_handler: bool) -> str:
        """Build prompt for Claude to translate VBA code.

        Args:
            vba_code: VBA code to translate
            _is_event_handler: Whether this is an event handler (reserved for future use)

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

        # Format mappings for prompt
        object_map_str = self._format_mappings_for_prompt(object_mappings, "Object Mappings")
        method_map_str = self._format_method_mappings(method_mappings)
        function_map_str = self._format_mappings_for_prompt(function_mappings, "Function Mappings")
        control_flow_str = self._format_control_flow(control_flow)
        declarations_str = self._format_control_flow(declarations)
        special_str = self._format_control_flow(special_cases)

        imports_str = "\n".join(required_imports)

        prompt = f"""Translate this VBA code to Python-UNO format for LibreOffice Calc.

VBA Code:
```vba
{vba_code}
```

Translation Rules:

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
