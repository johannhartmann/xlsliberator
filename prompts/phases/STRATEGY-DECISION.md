# Strategic Decision: Formula Translation Approach

**Date:** 2025-11-07
**Phase:** 6.2 (Formula Equivalence Validation)
**Status:** DECISION REQUIRED

## Problem Statement

Current manual formula translation implementation achieves only ~64% formula match rate, which is below the required 90% threshold (and far below the ideal 100%).

## Current Implementation

### What We Built:
1. **Excel Extraction** (`extract_excel.py`): Parse Excel files with openpyxl
2. **Formula Translation** (`formula_mapper.py`): Manual function mapping (SUM→SUMME, IF→WENN, etc.)
3. **ODS Building** (`write_ods.py`): Cell-by-cell construction via LibreOffice UNO API
4. **VBA Extraction** (`extract_vba.py`): Parse VBA modules from Excel
5. **VBA Translation** (`vba2py_uno.py` + `llm_vba_translator.py`): VBA→Python-UNO with LLM

### Achievements:
- ✅ Successfully converted 27,003 cells, 23,702 formulas
- ✅ 25 sheets processed
- ✅ VBA extraction working (29 modules)
- ✅ LLM-based VBA translation implemented with mapping injection
- ✅ Performance: 264s (~4.4 min) for real dataset
- ❌ Formula equivalence: ~64% (FAILS threshold)

### Root Cause of 64% Match Rate:
- IFERROR/MATCH formulas return incorrect values in LibreOffice
- Not a translation syntax issue - formulas are valid
- LibreOffice Calc calculates differently than Excel for certain patterns
- Example: `=IFERROR(MATCH($A2,D$2:D$19,0),"")` returns 0.0 instead of expected value

## Two Proposed Approaches

### Approach 1: Use LibreOffice Native Conversion

**Method:** Use LibreOffice's built-in converter: `soffice --headless --convert-to ods input.xlsm`

**Hypothesis:** LibreOffice's native conversion will achieve 100% formula equivalence because it uses the same calculation engine.

**Pros:**
- ✅ Likely 100% formula match rate (same engine)
- ✅ Much simpler codebase
- ✅ Faster conversion (single command)
- ✅ No manual formula translation needed
- ✅ Better formatting/style preservation

**Cons:**
- ❌ Less control over conversion process
- ❌ VBA is not converted (ODS doesn't support VBA)
- ❌ Need to verify named ranges, conditional formatting, etc.

**Implementation:**
```python
def convert_native(input_path: Path, output_path: Path) -> None:
    """Use LibreOffice native conversion."""
    subprocess.run([
        "soffice", "--headless", "--convert-to", "ods",
        str(input_path), "--outdir", str(output_path.parent)
    ])

    # Then add VBA translation on top
    extract_and_embed_vba(input_path, output_path)
```

**What We Keep:**
- VBA extraction and translation (our unique value)
- Macro embedding into converted ODS
- CLI/API interface
- Reporting

**What We Remove:**
- Manual formula translation (`formula_mapper.py`)
- Cell-by-cell ODS building (most of `write_ods.py`)
- Excel formula extraction (keep only for VBA)

---

### Approach 2: Fix Manual Translation

**Method:** Debug and fix our formula translation to achieve 100% match rate.

**Strategy:**
1. Identify all formula patterns causing mismatches
2. Add formula adaptation layer for LibreOffice quirks
3. Use LLM to translate complex/unsupported formulas
4. Extensive testing and validation

**Pros:**
- ✅ Full control over conversion
- ✅ Can customize for specific use cases
- ✅ Educational value (understand differences)
- ✅ Can add custom optimizations

**Cons:**
- ❌ Complex to maintain
- ❌ May never reach 100% (Excel/LibreOffice calculation differences)
- ❌ Time-consuming debugging
- ❌ Reinventing the wheel (LibreOffice already does this)

**Required Work:**
1. Debug IFERROR/MATCH issue (root cause: LibreOffice calc engine)
2. Add formula adaptation rules for known incompatibilities
3. Implement formula testing framework
4. Create comprehensive formula test suite
5. Iteratively fix mismatches

---

## Testing Plan

### Test 1: LibreOffice Native Conversion Quality

```bash
# Convert using native LibreOffice
soffice --headless --convert-to ods \
  tests/data/Bundesliga-Ergebnis-Tippspiel_V2.31_2025-26.xlsm \
  --outdir /tmp

# Test formula equivalence
python tools/test_equivalence.py \
  --excel tests/data/Bundesliga-Ergebnis-Tippspiel_V2.31_2025-26.xlsm \
  --ods /tmp/Bundesliga-Ergebnis-Tippspiel_V2.31_2025-26.ods \
  --tolerance 1e-9
```

**Expected Result:** ≥ 95% match rate (hopefully 100%)

### Test 2: VBA Preservation Check

```bash
# Check if VBA is preserved in native conversion
python -c "
from extract_vba import extract_vba_modules
modules = extract_vba_modules('/tmp/converted.ods')
print(f'VBA modules: {len(modules)}')
"
```

**Expected Result:** 0 modules (ODS doesn't support VBA natively)

### Test 3: Named Ranges & Features

```bash
# Verify named ranges preserved
python -c "
from write_ods import check_named_ranges
check_named_ranges('/tmp/converted.ods')
"
```

**Expected Result:** All named ranges present and functional

---

## Decision Criteria

| Criterion | Weight | Approach 1 (Native) | Approach 2 (Manual) |
|-----------|--------|---------------------|---------------------|
| Formula Equivalence (100% required) | ⭐⭐⭐⭐⭐ | 95-100% (likely) | 64% (current), uncertain max |
| VBA Translation (required) | ⭐⭐⭐⭐⭐ | Must add separately | Already implemented |
| Code Maintainability | ⭐⭐⭐⭐ | Simple, less code | Complex, more maintenance |
| Development Time | ⭐⭐⭐ | Fast (mostly done) | Slow (debugging needed) |
| Performance | ⭐⭐⭐ | Fast (native) | Medium (cell-by-cell) |
| Flexibility | ⭐⭐ | Less control | Full control |

**Critical Success Factor:** Must achieve 100% formula equivalence. If Approach 1 achieves this, it is the clear winner.

---

## Recommended Decision Process

1. **Test LibreOffice native conversion** (30 min)
   - Run conversion on test dataset
   - Measure formula equivalence
   - If ≥ 99%: Choose Approach 1
   - If < 95%: Need deeper analysis

2. **If Approach 1 succeeds:**
   - Implement hybrid: Native conversion + VBA translation
   - Test end-to-end pipeline
   - Document and deploy

3. **If Approach 1 fails:**
   - Investigate why native conversion doesn't achieve 100%
   - Evaluate if manual approach can surpass native
   - Consider hybrid: Native conversion + formula post-processing

---

## Next Actions

- [ ] Run Test 1: Native conversion formula equivalence
- [ ] Run Test 2: VBA preservation check
- [ ] Run Test 3: Named ranges verification
- [ ] Document test results
- [ ] Make decision based on test outcomes
- [ ] Update implementation plan
- [ ] Execute chosen approach

---

## Notes

- **Key Insight:** We may have been reinventing the wheel with manual formula translation
- **Focus:** Our unique value is VBA→Python-UNO translation, not formula conversion
- **Principle:** Use existing tools (LibreOffice native conversion) where they excel
- **Goal:** 100% formula equivalence is non-negotiable for production use
