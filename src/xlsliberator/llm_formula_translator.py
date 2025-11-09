"""LLM-based formula translation using Claude API."""

import json
import os
import re
from pathlib import Path
from typing import Any

import yaml
from anthropic import Anthropic
from loguru import logger


class LLMFormulaTranslator:
    """Translates Excel formulas to LibreOffice Calc using Claude LLM."""

    def __init__(
        self,
        cache_path: Path | None = None,
        incompatibility_rules_path: Path | None = None,
    ):
        """Initialize LLM translator.

        Args:
            cache_path: Optional path to cache translated formulas
            incompatibility_rules_path: Path to formula incompatibility rules YAML
        """
        self.client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        self.cache_path = cache_path or Path(".formula_cache.json")
        self.cache: dict[str, str] = self._load_cache()

        # Load incompatibility rules
        self.rules_path = incompatibility_rules_path or Path("rules/formula_incompatibilities.yaml")
        self.incompatibility_rules = self._load_incompatibility_rules()

    def _load_cache(self) -> dict[str, str]:
        """Load translation cache from disk."""
        if self.cache_path.exists():
            try:
                with open(self.cache_path) as f:
                    cache: dict[str, str] = json.load(f)
                    return cache
            except Exception as e:
                logger.warning(f"Failed to load formula cache: {e}")
        return {}

    def _save_cache(self) -> None:
        """Save translation cache to disk."""
        try:
            with open(self.cache_path, "w") as f:
                json.dump(self.cache, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save formula cache: {e}")

    def _load_incompatibility_rules(self) -> dict:
        """Load formula incompatibility rules from YAML."""
        if not self.rules_path.exists():
            logger.warning(f"Incompatibility rules not found: {self.rules_path}")
            return {}

        try:
            with open(self.rules_path) as f:
                rules: dict = yaml.safe_load(f)
                return rules
        except Exception as e:
            logger.warning(f"Failed to load incompatibility rules: {e}")
            return {}

    def translate_formula(
        self, excel_formula: str, locale: str = "en-US", rule_based_result: str | None = None
    ) -> str:
        """Translate an Excel formula to LibreOffice Calc format.

        This is a fallback translator - only called when rule-based translation fails.

        Args:
            excel_formula: Excel formula (including leading =)
            locale: Target locale (en-US, de-DE, etc.)
            rule_based_result: Result from rule-based translator (if available)

        Returns:
            Translated LibreOffice Calc formula
        """
        # Check cache first
        cache_key = f"{excel_formula}:{locale}"
        if cache_key in self.cache:
            logger.debug(f"LLM cache hit for formula: {excel_formula[:50]}...")
            return self.cache[cache_key]

        # Call Claude API for translation
        logger.info(f"LLM fallback for formula: {excel_formula[:50]}...")

        prompt = self._build_translation_prompt(excel_formula, locale)

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

            # Ensure formula starts with =
            if not translated.startswith("="):
                translated = "=" + translated

            # Cache the result
            self.cache[cache_key] = translated
            self._save_cache()

            logger.info(f"LLM translation: {excel_formula[:50]}... → {translated[:50]}...")
            return translated

        except Exception as e:
            logger.error(f"LLM translation failed: {e}")
            # Fallback: use rule-based result or original formula
            return rule_based_result if rule_based_result else excel_formula

    def _fix_offset_syntax(self, formula: str) -> str:
        """Fix LibreOffice OFFSET syntax issues that LLMs commonly generate.

        Problem: LLMs generate OFFSET(Sheet.$A$1;...) which causes #NAME? error
        Solution: Replace with OFFSET(Sheet.A1;...) - remove dollar signs from base ref

        Args:
            formula: Formula potentially containing broken OFFSET syntax

        Returns:
            Formula with fixed OFFSET syntax
        """

        # Pattern: OFFSET(SheetName.$A$1 or similar patterns with dollar signs
        # Match: OFFSET( followed by sheet name, dot, then $A$1 style reference
        # We need to remove the $ signs from the cell reference after the dot

        # Pattern matches: OFFSET(Sheet.$A$1... or OFFSET('Sheet-Name'.$A$1...
        pattern = r"(OFFSET\(['\"]?[\w\-]+['\"]?\.)(\$?)([A-Z]+)(\$?)(\d+)"

        def replace_dollars(match: re.Match[str]) -> str:
            """Remove dollar signs from the cell reference in OFFSET base."""
            prefix = match.group(1)  # OFFSET(Sheet.
            # groups 2,4 are the $ signs we want to remove
            col = match.group(3)  # Column letter (A, B, etc.)
            row = match.group(5)  # Row number
            return f"{prefix}{col}{row}"

        fixed = re.sub(pattern, replace_dollars, formula, flags=re.IGNORECASE)

        if fixed != formula:
            logger.debug("Post-processing: Fixed OFFSET syntax")
            logger.debug(f"  Before: {formula[:100]}...")
            logger.debug(f"  After:  {fixed[:100]}...")

        return fixed

    def translate_excel_to_calc(
        self,
        excel_formula: str,
        issue_type: str = "indirect_address_cross_sheet",
        sheet_name_mapping: dict[str, str] | None = None,
        custom_prompt: str | None = None,
    ) -> str:
        """Translate Excel formula with known incompatibility to Calc syntax.

        This method specifically fixes known incompatibilities between Excel and Calc,
        such as INDIRECT(ADDRESS(..., "SheetName")) patterns.

        Args:
            excel_formula: Excel formula with incompatibility
            issue_type: Type of incompatibility (key from rules YAML)
            sheet_name_mapping: Mapping of Excel sheet names to ODS quoted names
            custom_prompt: Optional custom prompt to use instead of rule-based prompt

        Returns:
            Calc-compatible formula

        Note:
            Uses specialized prompts based on incompatibility rules to ensure
            accurate translation. Results are cached for performance.
        """
        # Check cache first (skip caching if using custom prompt with retry logic)
        cache_key = f"repair:{issue_type}:{excel_formula}"
        if not custom_prompt and cache_key in self.cache:
            logger.debug(f"LLM cache hit for formula repair: {excel_formula[:50]}...")
            return self.cache[cache_key]

        # Use custom prompt if provided, otherwise build from rules
        if custom_prompt:
            prompt = custom_prompt
        else:
            # Get the rule for this issue type
            rule = self.incompatibility_rules.get(issue_type)
            if not rule:
                logger.warning(
                    f"No rule found for issue type: {issue_type}, falling back to original"
                )
                return excel_formula

            # Build specialized prompt from rule template
            prompt_template = rule.get("llm_prompt_template", "")
            if not prompt_template:
                logger.warning(f"No prompt template for issue type: {issue_type}")
                return excel_formula

            # Add sheet name mapping to prompt
            sheet_mapping_text = ""
            if sheet_name_mapping:
                sheet_mapping_text = (
                    "\n\nIMPORTANT - Sheet name mapping (Excel → LibreOffice Calc):\n"
                )
                for excel_name, calc_name in sheet_name_mapping.items():
                    sheet_mapping_text += f'  "{excel_name}" → {calc_name}\n'
                sheet_mapping_text += (
                    "\nUse the LibreOffice Calc names (right side) in your output."
                )

            prompt = prompt_template.format(excel_formula=excel_formula) + sheet_mapping_text

        # Call Claude API
        logger.info(f"LLM formula repair ({issue_type}): {excel_formula[:60]}...")

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

            # Ensure formula starts with =
            if not translated.startswith("="):
                translated = "=" + translated

            # Post-process: Fix LibreOffice OFFSET syntax issues
            # LLM often generates Sheet.$A$1 which causes #NAME? errors
            # LibreOffice Calc requires Sheet.A1 (no dollar signs in cross-sheet OFFSET)
            translated = self._fix_offset_syntax(translated)

            # Cache the result (only if not using custom prompt for retry)
            if not custom_prompt:
                self.cache[cache_key] = translated
                self._save_cache()

            logger.info(f"LLM repair: {excel_formula[:60]}... → {translated[:60]}...")
            return translated

        except Exception as e:
            logger.error(f"LLM formula repair failed: {e}")
            return excel_formula  # Return original if repair fails

    def _build_translation_prompt(self, excel_formula: str, locale: str) -> str:
        """Build prompt for Claude to translate formula.

        Args:
            excel_formula: Excel formula to translate
            locale: Target locale

        Returns:
            Prompt string for Claude
        """
        # Get locale-specific information
        locale_info: dict[str, dict[str, Any]] = {
            "en-US": {
                "separator": ",",
                "name": "English",
                "keep_english": True,
            },
            "de-DE": {
                "separator": ";",
                "name": "German",
                "function_examples": {
                    "SUM": "SUMME",
                    "IF": "WENN",
                    "MATCH": "VERGLEICH",
                    "VLOOKUP": "SVERWEIS",
                    "COUNT": "ANZAHL",
                },
            },
        }

        info: dict[str, Any] = locale_info.get(locale, locale_info["en-US"])

        if info.get("keep_english"):
            # For English, just return as-is
            return f"""Return this Excel formula exactly as provided, with no changes:

{excel_formula}

Output ONLY the formula, nothing else."""

        # For other locales, translate
        prompt = f"""Translate this Excel formula to LibreOffice Calc format for {info["name"]} locale ({locale}).

Excel Formula:
{excel_formula}

Requirements:
1. Convert function names to {info["name"]} equivalents (e.g., {", ".join(f"{k}→{v}" for k, v in info.get("function_examples", {}).items())})
2. Replace commas (,) with semicolons ({info["separator"]}) for function argument separators
3. Keep cell references and operators unchanged (A1, $B$2, +, -, *, /, etc.)
4. Preserve string literals, numbers, and logical values
5. Handle nested functions correctly
6. Keep the formula working identically to the original

Output ONLY the translated formula starting with =, no explanations or additional text.

Example:
Input: =IF(A1>0,SUM(B1:B10),0)
Output: =WENN(A1>0;SUMME(B1:B10);0)

Now translate the formula above:"""

        return prompt

    def translate_batch(self, formulas: list[str], locale: str = "en-US") -> list[str]:
        """Translate a batch of formulas.

        Args:
            formulas: List of Excel formulas
            locale: Target locale

        Returns:
            List of translated formulas
        """
        return [self.translate_formula(f, locale) for f in formulas]
