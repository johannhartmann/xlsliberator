"""Formula translation engine with locale support and tokenizer (Phase F5)."""

import enum
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml
from loguru import logger


class FormulaMappingError(Exception):
    """Raised when formula mapping fails."""


class TokenType(enum.Enum):
    """Token types for formula parsing."""

    EQUALS = "EQUALS"  # =
    FUNCTION = "FUNCTION"  # SUM, IF, VLOOKUP, etc.
    LPAREN = "LPAREN"  # (
    RPAREN = "RPAREN"  # )
    COMMA = "COMMA"  # ,
    SEMICOLON = "SEMICOLON"  # ;
    COLON = "COLON"  # :
    CELL_REF = "CELL_REF"  # A1, $B$2, etc.
    RANGE_REF = "RANGE_REF"  # A1:B10
    NUMBER = "NUMBER"  # 42, 3.14
    STRING = "STRING"  # "text"
    OPERATOR = "OPERATOR"  # +, -, *, /, ^, &, <, >, =
    WHITESPACE = "WHITESPACE"  # spaces
    UNKNOWN = "UNKNOWN"  # unrecognized


@dataclass
class Token:
    """Represents a token in a formula."""

    type: TokenType
    value: str
    position: int


class FormulaTokenizer:
    """Tokenizes Excel formulas for safe translation."""

    # Regex patterns for token recognition
    PATTERNS = [
        (TokenType.STRING, r'"(?:[^"\\]|\\.)*"'),  # String literals
        (TokenType.CELL_REF, r"\$?[A-Z]+\$?\d+"),  # Cell references (A1, $B$2)
        (TokenType.FUNCTION, r"\b[A-Z_][A-Z0-9_]*(?=\s*\()"),  # Function names
        (TokenType.NUMBER, r"\b\d+(?:\.\d+)?(?:[eE][+-]?\d+)?\b"),  # Numbers
        (TokenType.LPAREN, r"\("),
        (TokenType.RPAREN, r"\)"),
        (TokenType.COMMA, r","),
        (TokenType.SEMICOLON, r";"),
        (TokenType.COLON, r":"),
        (
            TokenType.OPERATOR,
            r"(?:<=|>=|<>|[+\-*/^&<>=])",
        ),  # Operators (multi-char first)
        (TokenType.EQUALS, r"="),
        (TokenType.WHITESPACE, r"\s+"),
    ]

    def __init__(self) -> None:
        """Initialize tokenizer with compiled patterns."""
        self.compiled_patterns = [
            (token_type, re.compile(pattern)) for token_type, pattern in self.PATTERNS
        ]

    def tokenize(self, formula: str) -> list[Token]:
        """Tokenize a formula string.

        Args:
            formula: Formula string (with or without leading =)

        Returns:
            List of tokens

        Raises:
            FormulaMappingError: If tokenization fails
        """
        tokens: list[Token] = []
        pos = 0
        formula_len = len(formula)

        while pos < formula_len:
            matched = False

            for token_type, pattern in self.compiled_patterns:
                match = pattern.match(formula, pos)
                if match:
                    value = match.group(0)
                    tokens.append(Token(type=token_type, value=value, position=pos))
                    pos = match.end()
                    matched = True
                    break

            if not matched:
                # Unknown token - take single character
                tokens.append(Token(type=TokenType.UNKNOWN, value=formula[pos], position=pos))
                pos += 1

        return tokens


# Global formula mapping (loaded from YAML)
_formula_mapping: dict[str, dict[str, Any]] | None = None
_locale_config: dict[str, dict[str, str]] | None = None


def _load_formula_mapping() -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, str]]]:
    """Load formula mapping from YAML file.

    Returns:
        Tuple of (formula_mapping, locale_config)

    Raises:
        FormulaMappingError: If YAML file cannot be loaded
    """
    global _formula_mapping, _locale_config

    if _formula_mapping is not None and _locale_config is not None:
        return _formula_mapping, _locale_config

    # Find YAML file
    yaml_path = Path(__file__).parent.parent.parent / "rules" / "formula_map.yaml"

    if not yaml_path.exists():
        raise FormulaMappingError(f"Formula mapping file not found: {yaml_path}")

    try:
        with open(yaml_path) as f:
            data = yaml.safe_load(f)

        # Extract function mappings (all keys except 'locales')
        _formula_mapping = {k: v for k, v in data.items() if k != "locales"}
        _locale_config = data.get("locales", {})

        logger.debug(
            f"Loaded {len(_formula_mapping)} function mappings and "
            f"{len(_locale_config)} locale configs from {yaml_path}"
        )

        return _formula_mapping, _locale_config

    except Exception as e:
        raise FormulaMappingError(f"Failed to load formula mapping: {e}") from e


def map_formula(formula: str, locale: str = "en-US") -> str:
    """Map Excel formula to LibreOffice Calc formula (Phase F5 - tokenizer-based).

    Args:
        formula: Excel formula string (e.g., "=SUM(A1:A10)")
        locale: Target locale ("en-US" or "de-DE")

    Returns:
        Mapped formula string with translated function names and separators

    Raises:
        FormulaMappingError: If mapping fails

    Note:
        Phase F5 implementation with proper tokenizer and locale-aware separator handling.
    """
    if not formula or not formula.startswith("="):
        logger.warning(f"Invalid formula format: {formula}")
        return formula

    # Load mapping tables
    func_mapping, locale_config = _load_formula_mapping()

    # Get locale-specific separator
    locale_sep = locale_config.get(locale, {}).get("separator", ",")

    # Tokenize formula
    tokenizer = FormulaTokenizer()
    tokens = tokenizer.tokenize(formula)

    # Translate tokens
    translated_tokens: list[str] = []

    for token in tokens:
        if token.type == TokenType.FUNCTION:
            # Translate function name
            func_upper = token.value.upper()
            if func_upper in func_mapping:
                # Get translated function name for locale
                translated = func_mapping[func_upper].get(locale, token.value)
                translated_tokens.append(translated)
            else:
                # Unknown function - keep as-is and log warning
                logger.warning(f"Unsupported function in formula: {token.value}")
                translated_tokens.append(token.value)

        elif token.type == TokenType.COMMA:
            # Replace comma with locale-specific separator
            translated_tokens.append(locale_sep)

        elif token.type == TokenType.WHITESPACE:
            # Preserve whitespace
            translated_tokens.append(token.value)

        else:
            # All other tokens (cell refs, numbers, strings, operators, etc.) - keep as-is
            translated_tokens.append(token.value)

    result = "".join(translated_tokens)
    logger.debug(f"Mapped formula: {formula} -> {result} (locale: {locale})")
    return result


def is_supported_formula(formula: str) -> bool:
    """Check if formula uses only supported functions.

    Args:
        formula: Excel formula string

    Returns:
        True if all functions are supported, False otherwise
    """
    if not formula or not formula.startswith("="):
        return False

    try:
        # Load mapping
        func_mapping, _ = _load_formula_mapping()

        # Tokenize formula
        tokenizer = FormulaTokenizer()
        tokens = tokenizer.tokenize(formula)

        # Check all function tokens
        for token in tokens:
            if token.type == TokenType.FUNCTION:
                func_upper = token.value.upper()
                if func_upper not in func_mapping:
                    logger.debug(f"Unsupported function: {token.value}")
                    return False

        return True

    except Exception as e:
        logger.warning(f"Error checking formula support: {e}")
        return False


def get_formula_functions(formula: str) -> set[str]:
    """Extract all function names from a formula.

    Args:
        formula: Excel formula string

    Returns:
        Set of function names (uppercase)
    """
    if not formula or not formula.startswith("="):
        return set()

    try:
        tokenizer = FormulaTokenizer()
        tokens = tokenizer.tokenize(formula)

        return {token.value.upper() for token in tokens if token.type == TokenType.FUNCTION}

    except Exception as e:
        logger.warning(f"Error extracting functions from formula: {e}")
        return set()
