"""LLM-based formula translation using Claude API."""

import json
import os
from pathlib import Path

from anthropic import Anthropic
from loguru import logger


class LLMFormulaTranslator:
    """Translates Excel formulas to LibreOffice Calc using Claude LLM."""

    def __init__(self, cache_path: Path | None = None):
        """Initialize LLM translator.

        Args:
            cache_path: Optional path to cache translated formulas
        """
        self.client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
        self.cache_path = cache_path or Path(".formula_cache.json")
        self.cache: dict[str, str] = self._load_cache()

    def _load_cache(self) -> dict[str, str]:
        """Load translation cache from disk."""
        if self.cache_path.exists():
            try:
                with open(self.cache_path) as f:
                    return json.load(f)
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

            translated = response.content[0].text.strip()

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

    def _build_translation_prompt(self, excel_formula: str, locale: str) -> str:
        """Build prompt for Claude to translate formula.

        Args:
            excel_formula: Excel formula to translate
            locale: Target locale

        Returns:
            Prompt string for Claude
        """
        # Get locale-specific information
        locale_info = {
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

        info = locale_info.get(locale, locale_info["en-US"])

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
