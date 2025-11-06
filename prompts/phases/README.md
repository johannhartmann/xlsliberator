# XLSLiberator Implementation Phases

This directory contains the split implementation prompts for the xlsliberator project. Use these prompts step-by-step while tracking progress in `../checklist.md`.

## Environment Setup
- **Python Environment:** conda environment `xlsliberator` (already activated)
- **Package Manager:** `uv` for fast dependency management and installation
- **Tooling:** `ruff` (formatting & linting), `mypy` (type checking), `pytest` (testing)
- **Project Structure:** Modern Python module with `pyproject.toml`

## Usage

Work through the phases in order:

1. **F0_project_kickoff.md** - Architecture & Feasibility Roadmap
2. **F1_repo_skeleton.md** - Repository Setup & Tooling
3. **F2_libreoffice_harness.md** - LibreOffice UNO Connection
4. **F3_excel_ingestion.md** - Excel File Parsing (all formats)
5. **F4_mini_ods_writer.md** - Basic ODS Writer
6. **F5_formula_mapper.md** - Formula Translation Engine
7. **F6_macro_embedding.md** - Python Macro Embedding
8. **F7_vba_extraction.md** - VBA Code Extraction
9. **F8_vba_translator.md** - VBA to Python Translation
10. **F9_tables_listobjects.md** - Excel Tables Support
11. **F10_charts_mvp.md** - Chart Conversion
12. **F11_formula_equivalence.md** - Formula Verification Tests
13. **F12_api_cli.md** - CLI & API Implementation
14. **F13_scorecard.md** - Automated Quality Gates
15. **F14_windows_validator.md** - Windows Excel COM Validation (optional)
16. **F15_performance_benchmarks.md** - Performance & Stability Tests
17. **F16_real_dataset.md** - Real-world Dataset Testing
18. **F17_fallback_path.md** - Fallback Import Strategy

## Quality Gates

Each phase has measurable quality gates defined. Track your progress in `../checklist.md` and verify:

- All tests pass before moving to the next phase
- Gates are documented in the conversion reports
- Any red gates are fixed before proceeding

## Security Notes

- VBA is analyzed **statically only** - no runtime execution
- Excel COM validator runs in **isolated Windows sandbox only**
- Path hardening and input validation throughout

## Commands Reference

```bash
# Install dependencies with uv
uv pip install -e ".[dev]"

# Run all tests
pytest -q

# Run specific phase tests
pytest tests/unit/test_<module>.py -q
pytest tests/it/test_<feature>.py -q

# Quality checks
make fmt && make lint && make typecheck && make test

# Or run individually
ruff format .
ruff check .
mypy src/
pytest

# Generate scorecard
python -m tools.scorecard out/report.json > out/feasibility_scorecard.md
```
