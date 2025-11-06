# XLSLiberator Quality Gates

This document defines all measurable quality gates for the xlsliberator project. Each gate must pass before proceeding to dependent milestones.

---

## Gate Status Legend

- 🟢 **Green:** Gate passed, criteria met
- 🟡 **Yellow:** Partial success, needs attention
- 🔴 **Red:** Gate failed, must fix before proceeding
- ⚪ **Not Started:** Gate not yet evaluated

---

## Foundation Gates (v0.1)

### G0: Project Setup
**Milestone:** M1 - Project Setup & Architecture
**Criteria:**
- Repository structure complete (`src/`, `tests/`, `rules/`, `docs/`)
- `pyproject.toml` with all dependencies
- All QA tools configured (ruff, mypy, pytest)

**Command:**
```bash
pytest && ruff check . && mypy src/
```

**Success Metrics:**
- Exit code 0 for all commands
- No configuration errors

**Status:** ⚪ Not Started

---

### G2: LibreOffice UNO Stability
**Milestone:** M2 - LibreOffice UNO Harness
**Criteria:**
- 10/10 connection cycles succeed
- No memory leaks detected
- Clean disconnect/cleanup

**Command:**
```bash
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &
pytest tests/it/test_uno_conn.py -q
```

**Success Metrics:**
- 10/10 `connect_lo()` → operations → close cycles pass
- Memory usage stable across cycles
- No zombie processes

**Status:** ⚪ Not Started

---

### G3: Excel Ingestion Completeness
**Milestone:** M3 - Excel Ingestion Engine
**Criteria:**
- ≥99% formulas extracted from test corpus
- All test files parse without errors
- IR serializes to JSON

**Command:**
```bash
pytest tests/unit/test_extract_excel.py -q
```

**Success Metrics:**
- Formula extraction rate: **≥99%**
- NamedRanges detected: **100%**
- Zero parse failures on test corpus

**Status:** ⚪ Not Started

---

### G4: ODS Recalc Correctness
**Milestone:** M4 - Minimal ODS Writer
**Criteria:**
- 10 core formulas calculate correctly
- Values match expected within tolerance

**Command:**
```bash
pytest tests/it/test_ods_writer_smoke.py -q
```

**Success Metrics:**
- 10/10 formulas produce correct values
- Tolerance: **±1e-9**
- No recalc errors

**Status:** ⚪ Not Started

---

### G5: Formula Mapper Accuracy
**Milestone:** M5 - Formula Mapper v1
**Criteria:**
- ≥90% syntactically correct translations
- Locale support verified (en-US, de-DE)

**Command:**
```bash
pytest tests/unit/test_formula_mapper.py -q
```

**Success Metrics:**
- Translation success rate: **≥90%**
- All existing IT tests still pass
- Both locales work correctly

**Status:** ⚪ Not Started

---

## Macros & Events Gates (v0.2)

### G6: Python Macro Embedding
**Milestone:** M6 - Python Macro Embedding
**Criteria:**
- Event fires exactly once
- Marker cell set correctly
- No crashes or double-calls

**Command:**
```bash
pytest tests/it/test_macro_embed.py -q
```

**Success Metrics:**
- `on_open` event fires: **exactly 1 time**
- Marker cell value correct
- No exceptions during open

**Status:** ⚪ Not Started

---

### G7: VBA Extraction Completeness
**Milestone:** M7 - VBA Extraction & Analysis
**Criteria:**
- 100% modules detected
- Dependency graph builds correctly
- Top API calls recognized

**Command:**
```bash
pytest tests/unit/test_extract_vba.py -q
```

**Success Metrics:**
- Module detection rate: **100%**
- Dependency graph acyclic and complete
- API token recognition: **≥95%** for common calls

**Status:** ⚪ Not Started

---

### G8: VBA Translation Functional
**Milestone:** M8 - VBA→Python Translator v1
**Criteria:**
- Translated handlers execute successfully
- Markers set correctly in E2E test

**Command:**
```bash
pytest tests/it/test_translated_macro_runs.py -q
```

**Success Metrics:**
- Translation compiles without syntax errors
- Translated handler runs and sets marker
- No runtime exceptions

**Status:** ⚪ Not Started

---

## Advanced Features Gates (v0.3)

### G9: Tables/ListObjects Support
**Milestone:** M9 - Tables/ListObjects Support
**Criteria:**
- ≥90% table formulas correct
- Structured references translate properly

**Command:**
```bash
pytest tests/it/test_tables_roundtrip.py -q
```

**Success Metrics:**
- Table formula correctness: **≥90%**
- AutoFilter preserved
- Named ranges created correctly

**Status:** ⚪ Not Started

---

### G10: Charts Conversion
**Milestone:** M10 - Charts MVP
**Criteria:**
- Series count matches original
- Titles and legends present

**Command:**
```bash
pytest tests/it/test_charts_basic.py -q
```

**Success Metrics:**
- Chart creation success: **≥80%**
- Series count accurate: **100%**
- Titles/legends preserved: **≥90%**

**Status:** ⚪ Not Started

---

### G11: Formula Equivalence
**Milestone:** M11 - Formula Equivalence Testing
**Criteria:**
- ≥95% values within tolerance band
- Outliers documented

**Command:**
```bash
pytest tests/it/test_formula_equivalence.py -q
```

**Success Metrics:**
- Value equivalence rate: **≥95%**
- Tolerance: **±1e-9**
- Outliers analyzed and reported

**Status:** ⚪ Not Started

---

## Integration Gates (v1.0)

### G12: CLI Smoke Test
**Milestone:** M12 - API & CLI
**Criteria:**
- CLI executes successfully
- Report generated and valid
- Exit code 0 on success

**Command:**
```bash
xlsliberator convert tests/data/sample.xlsm out/sample.ods --locale de-DE
```

**Success Metrics:**
- Command completes successfully
- ODS file created and opens in LibreOffice
- Report contains all expected sections

**Status:** ⚪ Not Started

---

### G13: Scorecard Generation
**Milestone:** M13 - Feasibility Scorecard
**Criteria:**
- Scorecard generates correctly
- All gates G0-G12 status accurate

**Command:**
```bash
python -m tools.scorecard out/report.json > out/feasibility_scorecard.md
```

**Success Metrics:**
- Scorecard file created
- Traffic light indicators match actual gate status
- No generation errors

**Status:** ⚪ Not Started

---

## Validation Gates (v1.1)

### G14: Windows Excel Validator (Optional)
**Milestone:** M14 - Windows Excel Validator
**Criteria:**
- Value deviation ≤1e-9 vs Excel COM
- Test passes on Windows, skips elsewhere

**Command:**
```bash
pytest -q -k win_excel_validator
```

**Success Metrics:**
- Value deviation: **≤1e-9**
- No COM errors on Windows
- Graceful skip on Linux

**Status:** ⚪ Not Started (Optional)

---

### G15: Performance & Stability
**Milestone:** M15 - Performance & Stability
**Criteria:**
- Benchmarks within targets
- 100/100 stability cycles pass

**Command:**
```bash
pytest tests/bench -q
pytest tests/it/test_lo_stability.py -q
```

**Success Metrics:**
- Stability cycles: **100/100** no crashes
- Ingestion: **≥50k cells/min**
- Formula mapping: **≥10k formulas/min**
- Memory: **<2GB peak** per file

**Status:** ⚪ Not Started

---

### G16: Real Dataset Success
**Milestone:** M16 - Real Dataset Testing
**Criteria:**
- ≥1 real file converts E2E successfully
- ≥80% tests green on first run

**Command:**
```bash
pytest tests/real/test_convert_real.py -q
```

**Success Metrics:**
- E2E success: **≥1 file** fully converts
- Test pass rate: **≥80%** on first run
- Formula equivalence: **≥90%** on real data

**Status:** ⚪ Not Started

---

## Robustness Gates (v1.2)

### G17: Fallback Path
**Milestone:** M17 - Fallback Import Path
**Criteria:**
- Fallback triggers when coverage low
- No hard failures, ODS always created

**Command:**
```bash
xlsliberator convert complex.xlsm out.ods --allow-fallback
```

**Success Metrics:**
- Fallback triggers correctly when needed
- ODS file always generated (if LO can import)
- Report marks fallback usage
- E2E tests remain green

**Status:** ⚪ Not Started

---

## Gate Dependency Graph

```
G0 (Setup)
  ↓
G2 (UNO) ──────┐
  ↓            ↓
G3 (Ingestion) G4 (ODS Writer)
  ↓            ↓
G7 (VBA Extract) G5 (Formula Mapper)
  ↓            ↓
G8 (VBA Translator) G6 (Macro Embed)
  ↓            ↓
  └─────┬──────┘
        ↓
  G9 (Tables) + G10 (Charts)
        ↓
  G11 (Equivalence)
        ↓
  G12 (CLI) ────→ G13 (Scorecard)
        ↓
  G14 (Windows) + G15 (Performance)
        ↓
  G16 (Real Data)
        ↓
  G17 (Fallback)
```

---

## Critical Path

**Must Pass Before v1.0:**
- G0, G2, G3, G4, G5, G6, G7, G8, G11, G12, G13

**Can Be Deferred:**
- G14 (Windows validator - optional)
- G9, G10 (Tables/Charts - can be v1.1)

**High Risk Gates:**
- G2 (UNO stability - foundational)
- G5 (Formula mapper - core value)
- G8 (VBA translator - complex)
- G11 (Equivalence - quality metric)

---

## Reporting

After each milestone:
1. Update this table with actual results
2. Generate scorecard: `python -m tools.scorecard`
3. Document any red gates with root cause
4. Update `prompts/checklist.md` with progress

---

## Current Status Summary

| Version | Total Gates | Passed | Failed | Not Started |
|---------|-------------|--------|--------|-------------|
| v0.1    | 5           | 0      | 0      | 5           |
| v0.2    | 3           | 0      | 0      | 3           |
| v0.3    | 3           | 0      | 0      | 3           |
| v1.0    | 2           | 0      | 0      | 2           |
| v1.1    | 3           | 0      | 0      | 3           |
| v1.2    | 1           | 0      | 0      | 1           |
| **Total** | **17**    | **0**  | **0**  | **17**      |

---

*Last Updated: 2025-11-06*
*Status: All gates defined, implementation not started*
