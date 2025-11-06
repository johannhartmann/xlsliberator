# XLSLiberator Feasibility Plan & Roadmap

## Executive Summary

**xlsliberator** is an Excel-to-LibreOffice Calc converter that transforms `.xlsx/.xlsm/.xlsb/.xls` files into fully functional `.ods` files with:
- Formula translation with locale support (de-DE/en-US)
- VBA-to-Python-UNO macro conversion
- Embedded Python macros with event handling
- Tables, Charts, and Forms preservation

This plan outlines a phased implementation approach with measurable feasibility gates at each step.

---

## Technology Stack

- **Language:** Python 3.11+
- **Package Manager:** uv (fast dependency management)
- **Excel Parsing:** openpyxl, pyxlsb, oletools
- **ODS Generation:** odfpy + LibreOffice UNO API
- **Code Quality:** ruff (format/lint), mypy (types), pytest (tests)
- **Environment:** conda environment `xlsliberator`

---

## Architecture Overview

### Data Flow

```
Excel File (.xlsx/.xlsm/.xlsb/.xls)
    ↓
[Excel Ingestion] → WorkbookIR (Intermediate Representation)
    ↓
[VBA Extraction] → VbaModuleIR + Dependency Graph
    ↓
[Formula Mapper] → Translated Formulas (locale-aware)
    ↓
[ODS Writer] → LibreOffice Calc Document (via UNO)
    ↓
[VBA→Python Translator] → Python-UNO Code
    ↓
[Macro Embedder] → Scripts/python/*.py in ODS
    ↓
[Tables/Charts/Forms] → Native Calc Features
    ↓
Final .ods File
```

### Key Components

1. **IR Models** - Pydantic-based intermediate representation
2. **Excel Extractors** - Format-specific parsers
3. **Formula Engine** - Tokenizer + rule-based mapper
4. **UNO Bridge** - LibreOffice headless connection
5. **VBA Translator** - AST-based code generator
6. **Feature Converters** - Tables, Charts, Forms

---

## Implementation Milestones

### v0.1 - Foundation (Milestones 1-5)

#### Milestone 1: Project Setup & Architecture
**Deliverables:**
- Repository structure (`src/`, `tests/`, `rules/`, `docs/`)
- `pyproject.toml` with dependencies
- Dev tooling (ruff, mypy, pytest)
- CI/CD configuration
- This feasibility plan

**Risks:** Tooling conflicts, dependency resolution
**Metrics:** All QA tools run successfully
**Gate G0:** `pytest && ruff check && mypy` pass

---

#### Milestone 2: LibreOffice UNO Harness
**Deliverables:**
- `uno_conn.py` - Connection management
- Helpers: `connect_lo()`, `new_calc()`, `save_as_ods()`, `recalc()`
- Integration tests with headless LibreOffice
- Connection stability tests

**Risks:** UNO API complexity, memory leaks, version compatibility
**Metrics:** 10/10 connect/disconnect cycles stable
**Gate G2:** `test_uno_conn.py` passes, no memory leaks

---

#### Milestone 3: Excel Ingestion Engine
**Deliverables:**
- `ir_models.py` - WorkbookIR, SheetIR, CellIR, NamedRangeIR
- `extract_excel.py` - Parser for .xlsx/.xlsm/.xlsb/.xls
- Formula extraction (all formats)
- NamedRanges, Charts metadata, ListObjects detection
- Unit tests with synthetic fixtures

**Risks:** Format coverage, formula edge cases, malformed files
**Metrics:** ≥99% formulas extracted from test corpus
**Gate G3:** All test files parse correctly, IR serializes to JSON

---

#### Milestone 4: Minimal ODS Writer
**Deliverables:**
- `write_ods.py` - IR to ODS converter (via UNO)
- Support for: sheets, cells, values, 10 core formulas
- Basic formula mapper (hardcoded IF/SUM/AVERAGE/VLOOKUP/etc.)
- Recalc smoke test

**Risks:** UNO API learning curve, formula syntax errors
**Metrics:** 10/10 test formulas calculate correctly
**Gate G4:** `recalc()` produces expected values (±1e-9 tolerance)

---

#### Milestone 5: Formula Mapper v1
**Deliverables:**
- `formula_mapper.py` - Tokenizer + rule engine
- `rules/formula_map.yaml` - ~25 core functions
- Locale support (`,` vs `;` for en-US/de-DE)
- Structured references detection (stub)
- Comprehensive unit tests

**Risks:** Parsing complexity, locale edge cases, function coverage
**Metrics:** ≥90% syntactically correct translations
**Gate G5:** Test suite ≥90% pass rate, existing IT tests still green

---

### v0.2 - Macros & Events (Milestones 6-8)

#### Milestone 6: Python Macro Embedding
**Deliverables:**
- `embed_macros.py` - Write Python modules into ODS
- `META-INF/manifest.xml` updates
- Event registration (on_open, sheet listeners)
- Integration test with marker cells

**Risks:** ODS structure complexity, event wiring
**Metrics:** Event fires exactly once per trigger
**Gate G6:** Marker cell set correctly, no crashes/double-calls

---

#### Milestone 7: VBA Extraction & Analysis
**Deliverables:**
- `extract_vba.py` - Parse vbaProject.bin
- VbaModuleIR models (Standard/Class/Form modules)
- Dependency graph builder
- API usage detection (Range/Cells/WorksheetFunction/etc.)
- Golden snippet tests

**Risks:** VBA dialect variations, binary format changes
**Metrics:** 100% modules detected, top API calls recognized
**Gate G7:** Dependency graph builds correctly, no parse failures

---

#### Milestone 8: VBA→Python Translator v1
**Deliverables:**
- `vba2py_uno.py` - Minimal VBA subset translator
- Support: Sub/Function, Dim, If/For/Select, Range/Cells calls
- `rules/vba_api_map.yaml` + `event_map.yaml`
- Golden translation tests
- E2E test: VBA snippet → Python → embedded → executes

**Risks:** Semantic gaps, API incompatibilities, edge cases
**Metrics:** Translated handlers run and set markers correctly
**Gate G8:** `test_translated_macro_runs.py` passes

---

### v0.3 - Advanced Features (Milestones 9-11)

#### Milestone 9: Tables/ListObjects Support
**Deliverables:**
- `tables_reader.py` - Extract Excel tables metadata
- `tables_to_uno.py` - Create Calc database ranges
- Structured reference translation in formula mapper
- Roundtrip test with table formulas

**Risks:** Structured reference complexity, AutoFilter quirks
**Metrics:** ≥90% table formulas correct
**Gate G9:** `test_tables_roundtrip.py` passes

---

#### Milestone 10: Charts MVP
**Deliverables:**
- `charts_reader.py` - Parse chart*.xml (Line/Column/Bar)
- `charts_to_uno.py` - Chart2 API usage
- Series, categories, titles, legends
- Integration test with visual validation (optional PNG export)

**Risks:** Chart type coverage, data range mapping
**Metrics:** Series count matches original, titles/legends present
**Gate G10:** `test_charts_basic.py` passes for 2+ chart types

---

#### Milestone 11: Formula Equivalence Testing
**Deliverables:**
- `testing_lo.py` - Recalc helpers + value comparison
- Sampling strategy (random + critical formulas)
- `test_formula_equivalence.py` - Excel cache vs Calc values
- Tolerance reporting (outliers documented)

**Risks:** Floating point differences, function behavior gaps
**Metrics:** ≥95% values within ±1e-9 tolerance
**Gate G11:** Equivalence tests pass, outliers analyzed

---

### v1.0 - Integration & CLI (Milestones 12-13)

#### Milestone 12: API & CLI
**Deliverables:**
- `api.py` - `convert(input, output, **options) -> Report`
- `cli.py` - Click-based CLI with progress indicators
- `report.py` - JSON/Markdown conversion reports
- Full pipeline integration
- CLI smoke tests

**Risks:** Error handling, user experience, edge cases
**Metrics:** Report complete and accurate, exit code 0 on success
**Gate G12:** `xlsliberator convert test.xlsm out.ods` succeeds

---

#### Milestone 13: Feasibility Scorecard
**Deliverables:**
- `tools/scorecard.py` - Automated gate status aggregator
- `feasibility_scorecard.md` generator (traffic light report)
- CI integration
- Snapshot tests

**Risks:** Metric interpretation, false positives/negatives
**Metrics:** All gates (G0-G12) correctly reflected
**Gate G13:** Scorecard generates and matches reality

---

### v1.1 - Validation & Performance (Milestones 14-16)

#### Milestone 14: Windows Excel Validator (Optional)
**Deliverables:**
- `test_win_excel_validator.py` - pywin32-based COM validator
- Excel CalculateFullRebuild comparison
- Sandbox execution (isolated Windows VM)
- Skip gracefully on Linux

**Risks:** COM API reliability, sandbox setup
**Metrics:** ≤1e-9 value deviation on sample cells
**Gate G14:** Test passes on Windows, skipped elsewhere

---

#### Milestone 15: Performance & Stability
**Deliverables:**
- `tests/bench/` - pytest-benchmark tests
- Metrics: cells/s, formulas/s, memory peak
- 100-cycle stability test (no leaks/crashes)
- Performance targets documented

**Risks:** Memory leaks, slow operations, resource exhaustion
**Metrics:** Benchmarks within targets, 100/100 stability
**Gate G15:** `test_lo_stability.py` completes successfully

---

#### Milestone 16: Real Dataset Testing
**Deliverables:**
- `tests/real/datasets.yaml` - Real Excel file registry
- `test_convert_real.py` - E2E validation suite
- Conversion reports with unsupported feature lists
- Scorecard updates

**Risks:** Real-world complexity, unexpected edge cases
**Metrics:** ≥80% tests green on first run, ≥1 file fully successful
**Gate G16:** At least one real dataset converts end-to-end

---

### v1.2 - Robustness (Milestone 17)

#### Milestone 17: Fallback Import Path
**Deliverables:**
- Fallback to LibreOffice native import when coverage low
- Post-processing: embed macros, fix NamedRanges
- Report fallback usage
- `--allow-fallback` CLI flag

**Risks:** Loss of control, quality degradation
**Metrics:** No hard failures, ODS always generated
**Gate G17:** Fallback triggers correctly, E2E remains green

---

## Testing Strategy

### Unit Tests
- Pure Python logic (IR models, formula mapper, VBA translator)
- Fast, no external dependencies
- 100+ tests across all modules
- Target: >80% code coverage

### Integration Tests (IT)
- Require LibreOffice headless
- Skippable via `LO_SKIP_IT=1`
- Test UNO operations, macro execution, recalc
- Target: All major workflows covered

### End-to-End Tests
- Full pipeline: Excel → ODS
- Real and synthetic datasets
- Performance benchmarks
- Target: ≥90% real datasets convert successfully

### Continuous Integration
- GitHub Actions / GitLab CI
- Matrix: Python 3.11/3.12, LibreOffice 7.x/24.x
- Quality gates on every commit
- Nightly real dataset runs

---

## Performance Targets

| Operation | Target | Measurement |
|-----------|--------|-------------|
| Excel Ingestion | ≥50k cells/min | Benchmark |
| Formula Mapping | ≥10k formulas/min | Benchmark |
| ODS Writing | ≥20k cells/min | Benchmark |
| Full Conversion | <5 min/file | Real datasets |
| Memory Peak | <2 GB/file | Real datasets |
| Stability | 100/100 cycles | Stress test |

---

## Locale Support

### Primary Locales
- **de-DE:** Semicolon separator (`;`), comma decimal
- **en-US:** Comma separator (`,`), period decimal

### Formula Examples
| Excel (en-US) | Calc (de-DE) |
|---------------|--------------|
| `=IF(A1>10, "Yes", "No")` | `=WENN(A1>10; "Yes"; "No")` |
| `=SUM(A1:A10)` | `=SUMME(A1:A10)` |
| `=VLOOKUP(...)` | `=SVERWEIS(...)` |

---

## Risk Register

### High Priority Risks

1. **UNO API Stability**
   - Mitigation: Robust error handling, connection retries, resource cleanup

2. **Formula Coverage Gaps**
   - Mitigation: Incremental mapping, unsupported function reporting, fallback path

3. **VBA Semantic Differences**
   - Mitigation: Focus on common patterns, document limitations, test extensively

4. **Performance at Scale**
   - Mitigation: Streaming where possible, memory profiling, optimization passes

5. **Real-World Excel Complexity**
   - Mitigation: Extensive real dataset testing, user feedback loop, iterative improvements

### Medium Priority Risks

6. **Locale Edge Cases**
   - Mitigation: Comprehensive locale test suite

7. **LibreOffice Version Compatibility**
   - Mitigation: Test against multiple LO versions

8. **Security (VBA Analysis)**
   - Mitigation: Static analysis only, no execution, sandboxing for validators

---

## Success Criteria

### Minimum Viable Product (MVP)
- ✅ Converts Excel to ODS with basic formulas
- ✅ ≥25 functions supported
- ✅ VBA macros translated (core subset)
- ✅ Tables and Charts preserved
- ✅ CLI tool available
- ✅ All gates G0-G12 green

### Version 1.0
- ✅ MVP criteria met
- ✅ ≥50 functions supported
- ✅ Real dataset success rate ≥80%
- ✅ Performance targets met
- ✅ Comprehensive documentation
- ✅ All gates G0-G16 green

### Version 1.x Enhancements
- Forms support (Phase 4.3)
- Extended VBA coverage
- Additional chart types
- Pivot table support
- Conditional formatting

---

## Go/No-Go Decision Points

After each milestone, evaluate:
1. **Gate Status:** All gates for this milestone green?
2. **Blocking Issues:** Any critical bugs preventing next phase?
3. **Risk Assessment:** Any new high-priority risks discovered?
4. **Timeline:** On track for overall delivery?

**Decision:** If ≥2 gates red OR critical blocker exists → STOP and fix before proceeding.

---

## Timeline Estimate

| Phase | Milestones | Estimated Duration |
|-------|------------|-------------------|
| v0.1 Foundation | M1-M5 | 2-3 weeks |
| v0.2 Macros | M6-M8 | 2-3 weeks |
| v0.3 Features | M9-M11 | 2 weeks |
| v1.0 Integration | M12-M13 | 1 week |
| v1.1 Validation | M14-M16 | 2 weeks |
| v1.2 Robustness | M17 | 1 week |
| **Total** | **17 milestones** | **10-14 weeks** |

*Note: Timeline assumes full-time dedicated development. Adjust based on actual availability.*

---

## Deliverables Checklist

- [x] This feasibility plan (`docs/feasibility_plan.md`)
- [ ] Quality gates table (`docs/gates.md`)
- [ ] Repository skeleton with modern tooling
- [ ] All 17 milestones implemented
- [ ] All quality gates passing
- [ ] Real dataset validation
- [ ] Production-ready CLI tool
- [ ] Comprehensive documentation

---

## Next Steps

1. **Review and Approve** this plan
2. **Setup Repository** (Milestone 1 / Prompt F1)
3. **Begin Implementation** following phase order F0→F17
4. **Track Progress** in `prompts/checklist.md`
5. **Generate Scorecards** after each milestone
6. **Iterate** based on real dataset feedback

---

*Document Version: 1.0*
*Last Updated: 2025-11-06*
*Status: Planning Complete - Ready for Implementation*
