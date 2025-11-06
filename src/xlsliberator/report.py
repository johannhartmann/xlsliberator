"""Conversion reporting module (Phase F12)."""

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path


@dataclass
class ConversionReport:
    """Report of Excel to ODS conversion."""

    input_file: str
    output_file: str
    success: bool

    # Cell statistics
    total_cells: int = 0
    total_formulas: int = 0
    formulas_translated: int = 0
    formulas_unsupported: int = 0

    # Named ranges
    named_ranges: int = 0

    # VBA/Macros
    vba_modules: int = 0
    vba_procedures: int = 0
    python_handlers: int = 0
    api_calls_mapped: int = 0

    # Sheets
    sheet_count: int = 0

    # Performance
    duration_seconds: float = 0.0

    # Issues
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    # Translation details
    locale: str = "en-US"

    def to_json(self) -> str:
        """Convert report to JSON string."""
        return json.dumps(asdict(self), indent=2)

    def to_markdown(self) -> str:
        """Convert report to Markdown string."""
        status = "✅ SUCCESS" if self.success else "❌ FAILED"

        md = f"""# Conversion Report

## Summary
- **Status**: {status}
- **Input**: `{self.input_file}`
- **Output**: `{self.output_file}`
- **Duration**: {self.duration_seconds:.2f}s
- **Locale**: {self.locale}

## Statistics

### Cells & Formulas
- Total cells: {self.total_cells:,}
- Total formulas: {self.total_formulas:,}
- Formulas translated: {self.formulas_translated:,} ({self._percentage(self.formulas_translated, self.total_formulas)})
- Formulas unsupported: {self.formulas_unsupported:,}

### Structure
- Sheets: {self.sheet_count}
- Named ranges: {self.named_ranges}

### VBA/Macros
- VBA modules: {self.vba_modules}
- VBA procedures: {self.vba_procedures}
- Python handlers created: {self.python_handlers}
- API calls mapped: {self.api_calls_mapped}

## Issues

### Warnings ({len(self.warnings)})
"""
        if self.warnings:
            for warning in self.warnings:
                md += f"- {warning}\n"
        else:
            md += "- None\n"

        md += f"\n### Errors ({len(self.errors)})\n"
        if self.errors:
            for error in self.errors:
                md += f"- {error}\n"
        else:
            md += "- None\n"

        return md

    def _percentage(self, value: int, total: int) -> str:
        """Calculate percentage string."""
        if total == 0:
            return "N/A"
        return f"{(value / total * 100):.1f}%"

    def save_json(self, path: str | Path) -> None:
        """Save report as JSON file."""
        Path(path).write_text(self.to_json())

    def save_markdown(self, path: str | Path) -> None:
        """Save report as Markdown file."""
        Path(path).write_text(self.to_markdown())
