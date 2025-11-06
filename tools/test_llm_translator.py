#!/usr/bin/env python3
"""Test LLM-based formula translator."""

import os

from loguru import logger

from xlsliberator.llm_formula_translator import LLMFormulaTranslator


def test_llm_translator() -> None:
    """Test formula translation with Claude."""
    if not os.environ.get("ANTHROPIC_API_KEY"):
        logger.error("ANTHROPIC_API_KEY not set")
        return

    translator = LLMFormulaTranslator()

    # Test formulas
    test_cases = [
        ("=SUM(A1:A10)", "en-US"),
        ("=IF(A1>0,SUM(B1:B10),0)", "en-US"),
        ('=IFERROR(MATCH($A2,D$2:D$19,0),"")', "en-US"),
        ("=SUM(A1:A10)", "de-DE"),
        ("=IF(A1>0,SUM(B1:B10),0)", "de-DE"),
        ('=IFERROR(MATCH($A2,D$2:D$19,0),"")', "de-DE"),
    ]

    logger.info("Testing LLM formula translation...")
    for formula, locale in test_cases:
        logger.info(f"\nInput ({locale}): {formula}")
        translated = translator.translate_formula(formula, locale)
        logger.info(f"Output: {translated}")


if __name__ == "__main__":
    test_llm_translator()
