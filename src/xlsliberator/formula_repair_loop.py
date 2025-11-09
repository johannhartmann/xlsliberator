"""Test-and-fix loop for formula repair with LLM retry logic."""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml
from loguru import logger

from xlsliberator.llm_formula_translator import LLMFormulaTranslator


@dataclass
class RepairAttempt:
    """Single repair attempt with result."""

    formula: str
    error_code: int
    error_name: str
    rule_book: str


@dataclass
class FormulaRepairResult:
    """Result of formula repair process."""

    original_excel: str
    original_ods: str
    final_formula: str | None
    success: bool
    attempts: list[RepairAttempt] = field(default_factory=list)
    error_history: list[str] = field(default_factory=list)


ERROR_CODES = {
    0: "No error",
    501: "#NULL!",
    502: "#DIV/0!",
    503: "#VALUE!",
    504: "#REF!",
    505: "#NAME?",
    506: "#NUM!",
    507: "#N/A",
    508: "Err:508 (Missing parenthesis)",
    509: "Err:509 (Missing operator)",
    510: "Err:510 (Missing variable)",
    511: "Err:511 (Missing separator)",
    512: "Err:512 (Formula overflow)",
    513: "Err:513 (String overflow)",
    514: "Err:514 (Internal overflow)",
    515: "Err:515 (Internal syntax error)",
    516: "Err:516 (Internal syntax error)",
    517: "Err:517 (Internal syntax error)",
    518: "Err:518 (Internal syntax error)",
    519: "Err:519 (No code)",
    520: "Err:520 (No value)",
    521: "Err:521 (No macro)",
    522: "Err:522 (Circular reference)",
    523: "Err:523 (Calculation)",
    524: "#REF! (Invalid reference)",
    525: "#NAME? (Unknown name)",
    526: "Err:526 (Invalid separator)",
    527: "Err:527 (Invalid format)",
}


class FormulaRepairLoop:
    """Test-and-fix loop for formula repair."""

    def __init__(
        self,
        rule_books_dir: Path | None = None,
        max_attempts: int = 10,
    ):
        """Initialize repair loop.

        Args:
            rule_books_dir: Directory containing transformation rule YAML files
            max_attempts: Maximum repair attempts per formula (default 10)
        """
        self.rule_books_dir = rule_books_dir or Path("rules/calc_transforms")
        self.max_attempts = max_attempts
        self.translator = LLMFormulaTranslator()
        self.rule_books = self._load_rule_books()

    def _load_rule_books(self) -> list[dict]:
        """Load all transformation rule books from directory."""
        if not self.rule_books_dir.exists():
            logger.warning(f"Rule books directory not found: {self.rule_books_dir}")
            return []

        rule_books = []
        for yaml_file in sorted(self.rule_books_dir.glob("*.yaml")):
            try:
                with open(yaml_file) as f:
                    rule_book = yaml.safe_load(f)
                    rule_book["_file"] = yaml_file.name
                    rule_books.append(rule_book)
                    logger.debug(f"Loaded rule book: {rule_book.get('name', yaml_file.name)}")
            except Exception as e:
                logger.warning(f"Failed to load rule book {yaml_file}: {e}")

        # Sort by priority
        rule_books.sort(key=lambda r: r.get("priority", 999))
        logger.info(f"Loaded {len(rule_books)} transformation rule books")
        return rule_books

    def _test_formula(self, doc: Any, sheet: Any, cell_addr: str, formula: str) -> tuple[int, str]:
        """Test formula in LibreOffice and return error code.

        Args:
            doc: UNO document object
            sheet: UNO sheet object
            cell_addr: Cell address like "B3"
            formula: Formula to test

        Returns:
            (error_code, error_name)
        """
        # Parse cell address
        col_str = "".join(c for c in cell_addr if c.isalpha())
        row_str = "".join(c for c in cell_addr if c.isdigit())

        # Convert column letters to index
        col_idx = 0
        for char in col_str:
            col_idx = col_idx * 26 + (ord(char.upper()) - ord("A") + 1)
        col_idx -= 1  # 0-indexed

        row_idx = int(row_str) - 1  # 0-indexed

        # Get cell and set formula
        cell = sheet.getCellByPosition(col_idx, row_idx)
        cell.setFormula(formula)

        # Recalculate
        doc.calculateAll()

        # Get error
        error_code = cell.getError()
        error_name = ERROR_CODES.get(error_code, f"Unknown error {error_code}")

        return error_code, error_name

    def _select_rule_book(self, excel_formula: str, attempt_count: int) -> dict | None:
        """Select appropriate rule book for formula.

        Args:
            excel_formula: Original Excel formula
            attempt_count: Current attempt number

        Returns:
            Selected rule book dict or None
        """
        # On later attempts, try fallback rule books
        if attempt_count > 3:
            # Try fallback/offset strategies
            fallback_books = [r for r in self.rule_books if "fallback" in r.get("name", "").lower()]
            if fallback_books:
                return fallback_books[0]

        # Check patterns
        for rule_book in self.rule_books:
            pattern_info = rule_book.get("pattern", {})
            regex = pattern_info.get("regex")

            if regex and regex != ".*":  # Skip catch-all patterns initially
                import re

                if re.search(regex, excel_formula, re.IGNORECASE):
                    # Check exclude pattern
                    exclude_regex = pattern_info.get("exclude_regex")
                    if exclude_regex and re.search(exclude_regex, excel_formula, re.IGNORECASE):
                        continue

                    return rule_book

        # Default: use first rule book
        return self.rule_books[0] if self.rule_books else None

    def _build_prompt(
        self,
        excel_formula: str,
        rule_book: dict,
        sheet_mapping: dict[str, str],
        attempts: list[RepairAttempt],
    ) -> str:
        """Build LLM prompt from rule book template.

        Args:
            excel_formula: Original Excel formula
            rule_book: Selected rule book
            sheet_mapping: Excel→ODS sheet name mapping
            attempts: Previous repair attempts

        Returns:
            Complete prompt string
        """
        prompt_template = rule_book.get("llm_prompt", "")

        # Build sheet mapping text
        sheet_mapping_text = "\n".join(
            f'  "{excel}" → {calc}' for excel, calc in sheet_mapping.items()
        )

        # Build attempt history
        attempt_history = ""
        if attempts:
            attempt_history = "Previous failed attempts:\n"
            for i, attempt in enumerate(attempts, 1):
                attempt_history += f"  Attempt {i} ({attempt.rule_book}): {attempt.formula}\n"
                attempt_history += f"    Error: {attempt.error_code} ({attempt.error_name})\n"

        # Get last error info
        error_code = attempts[-1].error_code if attempts else 0
        error_message = attempts[-1].error_name if attempts else "No error yet"

        # Fill template
        prompt: str = prompt_template.format(
            excel_formula=excel_formula,
            sheet_mapping=sheet_mapping_text,
            attempt_history=attempt_history,
            error_code=error_code,
            error_message=error_message,
        )

        return prompt

    def repair_formula(
        self,
        doc: Any,
        sheet: Any,
        cell_addr: str,
        excel_formula: str,
        ods_formula: str,
        sheet_mapping: dict[str, str],
    ) -> FormulaRepairResult:
        """Repair formula with test-and-fix loop.

        Args:
            doc: UNO document object
            sheet: UNO sheet object
            cell_addr: Cell address like "B3"
            excel_formula: Original Excel formula
            ods_formula: Current ODS formula (may be broken)
            sheet_mapping: Excel→ODS sheet name mapping

        Returns:
            FormulaRepairResult with outcome
        """
        result = FormulaRepairResult(
            original_excel=excel_formula,
            original_ods=ods_formula,
            final_formula=None,
            success=False,
        )

        # Test original ODS formula first
        error_code, error_name = self._test_formula(doc, sheet, cell_addr, ods_formula)

        if error_code == 0:
            # Already works!
            result.final_formula = ods_formula
            result.success = True
            return result

        # Record initial error
        result.error_history.append(f"Original ODS: {error_code} ({error_name})")

        # Repair loop
        for attempt_num in range(1, self.max_attempts + 1):
            logger.debug(f"Repair attempt {attempt_num}/{self.max_attempts} for {cell_addr}")

            # Select rule book
            rule_book = self._select_rule_book(excel_formula, attempt_num)
            if not rule_book:
                logger.warning("No rule book available")
                break

            rule_name = rule_book.get("name", "Unknown")

            # Build prompt with history
            prompt = self._build_prompt(excel_formula, rule_book, sheet_mapping, result.attempts)

            # Call LLM
            try:
                # Convert Excel syntax (commas) to ODS syntax for prompt
                excel_style = excel_formula.replace(";", ",")

                repaired = self.translator.translate_excel_to_calc(
                    excel_style,
                    issue_type="formula_repair_loop",
                    sheet_name_mapping=sheet_mapping,
                    custom_prompt=prompt,
                )

                # Convert back to LibreOffice semicolons
                repaired_ods = repaired.replace(",", ";")

                # Clean formula: Remove dollar signs from OFFSET base references
                # OFFSET(Sheet.$A$1,...) → OFFSET(Sheet.A1,...)
                import re

                repaired_ods = re.sub(
                    r"OFFSET\(([^.]+)\.\$([A-Z]+)\$(\d+)",  # Match OFFSET(Sheet.$A$1
                    r"OFFSET(\1.\2\3",  # Replace with OFFSET(Sheet.A1
                    repaired_ods,
                    flags=re.IGNORECASE,
                )

            except Exception as e:
                logger.error(f"LLM translation failed: {e}")
                result.error_history.append(f"Attempt {attempt_num}: LLM error: {e}")
                continue

            # Test repaired formula
            error_code, error_name = self._test_formula(doc, sheet, cell_addr, repaired_ods)

            # Record attempt
            attempt = RepairAttempt(
                formula=repaired_ods,
                error_code=error_code,
                error_name=error_name,
                rule_book=rule_name,
            )
            result.attempts.append(attempt)
            result.error_history.append(
                f"Attempt {attempt_num} ({rule_name}): {error_code} ({error_name})"
            )

            if error_code == 0:
                # Success!
                result.final_formula = repaired_ods
                result.success = True
                logger.success(f"Repaired {cell_addr} after {attempt_num} attempts")
                return result

            logger.debug(f"Attempt {attempt_num} failed: {error_name}")

        # Max attempts reached
        logger.warning(f"Failed to repair {cell_addr} after {self.max_attempts} attempts")
        return result
