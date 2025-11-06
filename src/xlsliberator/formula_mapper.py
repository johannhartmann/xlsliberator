"""Formula translation engine with locale support."""

from loguru import logger


class FormulaMappingError(Exception):
    """Raised when formula mapping fails."""


# Hardcoded mapping for Phase F4 (v0 - 10 core functions)
# This will be extended in Phase F5 with proper tokenizer and YAML rules
FORMULA_MAP_EN_US = {
    "IF": "IF",
    "SUM": "SUM",
    "AVERAGE": "AVERAGE",
    "SUMIF": "SUMIF",
    "SUMIFS": "SUMIFS",
    "COUNTIF": "COUNTIF",
    "COUNTIFS": "COUNTIFS",
    "INDEX": "INDEX",
    "MATCH": "MATCH",
    "VLOOKUP": "VLOOKUP",
}

FORMULA_MAP_DE_DE = {
    "IF": "WENN",
    "SUM": "SUMME",
    "AVERAGE": "MITTELWERT",
    "SUMIF": "SUMMEWENN",
    "SUMIFS": "SUMMEWENNS",
    "COUNTIF": "ZÄHLENWENN",
    "COUNTIFS": "ZÄHLENWENNS",
    "INDEX": "INDEX",
    "MATCH": "VERGLEICH",
    "VLOOKUP": "SVERWEIS",
}


def map_formula(formula: str, locale: str = "en-US") -> str:
    """Map Excel formula to LibreOffice Calc formula (v0 - basic implementation).

    Args:
        formula: Excel formula string (e.g., "=SUM(A1:A10)")
        locale: Target locale ("en-US" or "de-DE")

    Returns:
        Mapped formula string

    Raises:
        FormulaMappingError: If mapping fails

    Note:
        This is a minimal v0 implementation for Phase F4.
        Full tokenizer and locale-aware mapping will be implemented in Phase F5.
    """
    if not formula or not formula.startswith("="):
        logger.warning(f"Invalid formula format: {formula}")
        return formula

    # Get mapping table
    mapping = FORMULA_MAP_DE_DE if locale == "de-DE" else FORMULA_MAP_EN_US

    # Simple string replacement for now (v0)
    # This is a naive approach that will be replaced with proper tokenization in F5
    mapped = formula
    for excel_func, calc_func in mapping.items():
        # Replace function names (case-insensitive)
        # This is simplistic and doesn't handle function names inside strings
        import re

        pattern = rf"\b{excel_func}\b"
        mapped = re.sub(pattern, calc_func, mapped, flags=re.IGNORECASE)

    # Handle separator differences for de-DE
    if locale == "de-DE":
        # In German locale, use semicolon instead of comma for function arguments
        # This is a simplified approach - proper implementation in F5
        # For now, just note it for formulas that need it
        pass

    logger.debug(f"Mapped formula: {formula} -> {mapped} (locale: {locale})")
    return mapped


def is_supported_formula(formula: str) -> bool:
    """Check if formula uses only supported functions (v0).

    Args:
        formula: Excel formula string

    Returns:
        True if all functions are supported, False otherwise
    """
    if not formula or not formula.startswith("="):
        return False

    # Extract function names (simple approach for v0)
    import re

    # Find all function names (words followed by opening parenthesis)
    functions = re.findall(r"\b([A-Z_]+)\s*\(", formula, re.IGNORECASE)

    # Check if all are in our mapping
    supported = set(FORMULA_MAP_EN_US.keys())
    formula_funcs = {f.upper() for f in functions}

    unsupported = formula_funcs - supported
    if unsupported:
        logger.debug(f"Unsupported functions in formula: {unsupported}")
        return False

    return True
