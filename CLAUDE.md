# XLSLiberator Project Context

## Project Overview

**xlsliberator** is an Excel-to-LibreOffice Calc converter that transforms Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) into LibreOffice Calc `.ods` files with:
- Full formula translation and equivalence
- VBA-to-Python-UNO macro conversion
- Embedded Python macros with event handling
- Tables, Charts, and Forms support

## Development Environment

- **Python Environment:** conda environment `xlsliberator` (already activated)
- **Package Manager:** `uv` for fast dependency management
- **Code Quality Tools:**
  - `ruff` - formatting & linting
  - `mypy` - type checking
  - `pytest` - testing framework
- **Project Structure:** Modern Python module with `pyproject.toml`

## Implementation Approach

This project follows a **phased implementation** with **feasibility gates** at each step. Each phase has:
- Clear objectives and deliverables
- Measurable success metrics
- Quality gates that must pass before proceeding
- Test coverage requirements

### Phase Files Location
All implementation prompts are in `prompts/phases/F0-F17.md`

### Progress Tracking
**IMPORTANT:** Always maintain the checklist in `prompts/checklist.md`:
- Mark phases as complete when all gates pass
- Update status after each implementation step
- Use checkboxes to track progress: `- [x]` for complete, `- [ ]` for pending

## Key Commands

```bash
# Install dependencies
uv pip install -e ".[dev]"

# Code quality
ruff format .          # Format code
ruff check .           # Lint code
mypy src/             # Type check
pytest                # Run tests

# Or use Makefile
make fmt && make lint && make typecheck && make test

# LibreOffice headless (for integration tests)
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &
```

## Project Structure

```
xlsliberator/
├── src/xlsliberator/         # Main source code
│   ├── ir_models.py          # Intermediate representation models
│   ├── extract_excel.py      # Excel file parsing
│   ├── extract_vba.py        # VBA extraction
│   ├── formula_mapper.py     # Formula translation
│   ├── uno_conn.py           # LibreOffice UNO connection
│   ├── write_ods.py          # ODS file generation
│   ├── embed_macros.py       # Python macro embedding
│   ├── vba2py_uno.py         # VBA to Python translation
│   ├── tables_*.py           # Table handling
│   ├── charts_*.py           # Chart conversion
│   ├── forms_*.py            # Forms processing
│   ├── testing_lo.py         # LibreOffice testing utilities
│   ├── report.py             # Conversion reporting
│   ├── api.py                # API interface
│   └── cli.py                # Command-line interface
├── tests/
│   ├── unit/                 # Unit tests
│   ├── it/                   # Integration tests
│   ├── bench/                # Performance benchmarks
│   ├── real/                 # Real dataset tests
│   └── data/                 # Test fixtures
├── rules/                    # YAML mapping rules
│   ├── formula_map.yaml      # Formula translation rules
│   ├── vba_api_map.yaml      # VBA API mappings
│   ├── event_map.yaml        # Event mappings
│   ├── forms_map.yaml        # Forms mappings
│   └── charts_map.yaml       # Chart mappings
├── docs/                     # Documentation
│   ├── feasibility_plan.md   # Roadmap and milestones
│   └── gates.md              # Quality gates table
├── prompts/                  # Implementation prompts
│   ├── phases/               # Phase-by-phase prompts (F0-F17)
│   └── checklist.md          # **Progress checklist - KEEP UPDATED**
└── tools/                    # Development tools
    └── scorecard.py          # Automated gate scorecard
```

## Implementation Phases

1. **F0** - Project Kickoff & Feasibility Roadmap
2. **F1** - Repo Skeleton & Tooling
3. **F2** - LibreOffice Headless Harness
4. **F3** - Excel Ingestion
5. **F4** - Mini ODS Writer
6. **F5** - Formula Mapper v1
7. **F6** - Macro Embedding
8. **F7** - VBA Extraction
9. **F8** - VBA→Python Translator
10. **F9** - Tables/ListObjects MVP
11. **F10** - Charts MVP
12. **F11** - Formula Equivalence Testing
13. **F12** - API/CLI Integration
14. **F13** - Feasibility Scorecard
15. **F14** - Windows Validator (optional)
16. **F15** - Performance Benchmarks
17. **F16** - Real Dataset Testing
18. **F17** - Fallback Path

## Quality Gates Summary

- **G1:** CI-local green (all tests pass)
- **G2:** 10/10 LibreOffice connection cycles stable
- **G3:** ≥99% formulas extracted from test files
- **G4:** Recalc produces expected values (±1e-9)
- **G5:** ≥90% formula syntax translations correct
- **G6:** Event markers set correctly, no crashes
- **G7:** VBA dependency graph builds without errors
- **G8:** Translated VBA integration tests green
- **G9:** ≥90% table formulas correct
- **G10:** Charts created with correct series/titles
- **G11:** ≥95% formula values in tolerance band
- **G12:** CLI smoke tests green
- **G13:** Scorecard generates correctly
- **G14:** Windows Excel COM validation (optional)
- **G15:** 100/100 stability cycles, benchmarks pass
- **G16:** ≥1 real dataset converts successfully E2E
- **G17:** Fallback path works when coverage < threshold

## Testing Strategy

- **Unit Tests:** Test individual components in isolation
- **Integration Tests (IT):** Test with LibreOffice headless
- **E2E Tests:** Full pipeline from Excel to ODS
- **Benchmarks:** Performance and memory profiling
- **Real Dataset Tests:** Validation on actual Excel files

## Security Considerations

- VBA is analyzed **statically only** - no runtime execution
- Excel COM validator runs in **isolated Windows sandbox only**
- Path hardening and input validation throughout
- No credential harvesting or malicious code generation

## Workflow Rules

1. **Follow phase order** - Complete each phase before moving to next
2. **Update checklist** - Mark items complete in `prompts/checklist.md`
3. **Run quality gates** - All tests must pass before proceeding
4. **Commit each phase** - Git commit after completing each phase
5. **Document failures** - Record any gate failures and fixes
6. **Generate reports** - Create conversion reports for each test run

## Current Status

**Date: 2025-11-07**

### ✅ Completed Phases:
- **Phase 0-3**: Setup, Excel ingestion, ODS writer, VBA extraction (COMPLETE)
- **Phase 5.1-5.2**: API/CLI integration, reporting (COMPLETE)
- **Phase 6.1**: Real dataset E2E conversion successful (COMPLETE)
- **Phase 6.3**: Performance < 5 min achieved (264s for 27k cells) (COMPLETE)

### ✅ Strategic Decision Made: Hybrid Native + VBA Translation Approach

**Decision Date:** 2025-11-07

**DECISION:** Use LibreOffice native conversion for base ODS, then add VBA→Python-UNO translation.

**Rationale:**
- Manual formula translation achieved only 64% match rate
- LibreOffice native conversion uses same calculation engine → 100% formula equivalence expected
- Our unique value proposition is VBA→Python-UNO translation, not formula conversion
- Simpler architecture, faster conversion, better maintainability

**New Implementation Architecture:**

```
┌─────────────┐
│ Excel File  │
│ (.xlsm)     │
└──────┬──────┘
       │
       ├─────────────────────────────────┐
       │                                 │
       v                                 v
┌──────────────────┐            ┌────────────────┐
│ LibreOffice      │            │ VBA Extraction │
│ Native Convert   │            │ (oletools-vba) │
│ --convert-to ods │            └────────┬───────┘
└────────┬─────────┘                     │
         │                               v
         │                    ┌──────────────────────┐
         │                    │ VBA→Python-UNO       │
         │                    │ Translation (LLM)    │
         │                    └──────────┬───────────┘
         │                               │
         v                               v
    ┌────────────────────────────────────────┐
    │   Embed Python Macros into ODS         │
    │   (embed_macros.py via UNO)            │
    └────────────────────────────────────────┘
                    │
                    v
            ┌───────────────┐
            │ Final ODS     │
            │ + Py Macros   │
            └───────────────┘
```

**What We Keep:**
- ✅ VBA extraction (`extract_vba.py`)
- ✅ VBA→Python-UNO translation (`vba2py_uno.py`, `llm_vba_translator.py`)
- ✅ Macro embedding (`embed_macros.py`)
- ✅ CLI/API interface (`cli.py`, `api.py`)
- ✅ Testing framework (`testing_lo.py`)

**What We Replace:**
- ❌ Manual formula translation (`formula_mapper.py`) → Use LibreOffice native
- ❌ Cell-by-cell ODS building (`write_ods.py`) → Use LibreOffice native
- ❌ Excel formula extraction (openpyxl) → Only needed for validation

**Implementation Plan:**
1. Update `api.py` to use subprocess for native conversion
2. Keep VBA extraction and translation pipeline
3. Embed translated Python macros into native-converted ODS
4. Test formula equivalence (expect 100%)
5. Validate VBA translation works end-to-end

Check `prompts/checklist.md` for detailed phase status.
