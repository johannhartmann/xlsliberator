# Remaining Formula Issues (12% Mismatches)

**Current Status:** 88.21% formula equivalence (20,908/23,702 formulas matching)
**Remaining:** 2,794 mismatches (11.79%)

## Issue Categories

### 1. INDIRECT/ADDRESS Cross-Sheet References (Primary Issue)
**Count:** ~840 direct mismatches + ~19,356 cascading errors = **~20,196 total**

**Root Cause:** Excel's `ADDRESS()` function accepts sheet names as 5th parameter, but LibreOffice Calc does not support this syntax.

**Example Formula:**
```excel
=IF(INDIRECT(ADDRESS(48+COLUMN(),12,1,1,"Spielplan"))="ja",INDIRECT(ADDRESS(IF(MOD(B$1,2)=0,25,12),(ROW()-2)*2+4,1,1,"1-2")),"")
```

**Excel Behavior:**
- `ADDRESS(row, col, abs, a1, "Spielplan")` → `"Spielplan!$L$48"`
- `INDIRECT("Spielplan!$L$48")` → Value from that cell
- Result: Correct value (e.g., 0, text, number)

**LibreOffice Calc Behavior:**
- `ADDRESS(row, col, abs, a1, "Spielplan")` → `#NAME?` error (sheet name parameter not supported)
- All dependent formulas fail
- Result: `#NAME?` error cascades through dependent cells

**AI Repair Strategy:**
```
Pattern: ADDRESS(..., ..., ..., ..., "SheetName")
Replace: Convert to direct sheet reference syntax

Excel:    =INDIRECT(ADDRESS(row, col, 1, 1, "Sheet"))
LibreOffice: =INDIRECT("Sheet." & ADDRESS(row, col, 1, 1))

OR use OFFSET with sheet-qualified range:
  =OFFSET(Sheet.$A$1, row-1, col-1, 1, 1)
```

**Sample Affected Cells:**
- `Spieler!B3:K3` - All player data cells using INDIRECT/ADDRESS
- Cascades to `Spieler!AJ3` (SUM of B3:AI3 → all #NAME?)
- Affects all player sheets that reference match data dynamically

---

### 2. Text Formula Return Values
**Count:** 1 mismatch

**Issue:** Formula returns text in Excel, but Calc returns 0.0

**Example:**
- Cell: `Rangliste!B27`
- Excel: `"                                                    "` (spaces)
- Calc: `0.0`

**AI Repair Strategy:**
```
Check formula type, if should return text:
- Ensure TEXT() or string concatenation is explicit
- Wrap numeric results in TEXT() if needed
```

---

### 3. Other Mismatches
**Count:** 347 mismatches

**Likely Causes:**
- Cascading errors from INDIRECT/ADDRESS issues
- Cell type mismatches (formula result stored as different type)
- Minor calculation engine differences

**AI Repair Strategy:**
```
Analyze each mismatch individually:
1. Check if it's a dependent cell of INDIRECT/ADDRESS error
2. Check for Excel-specific functions without Calc equivalent
3. Check for formula syntax differences
```

---

## Repair Priority

### High Priority (Fixes ~98% of issues)
**Fix INDIRECT/ADDRESS cross-sheet references**
- 840 source formulas need rewriting
- Will resolve ~19,356 cascading #NAME? errors
- Impact: 88.21% → ~99.8% formula equivalence

### Low Priority
- Text formula issues (1 formula)
- Other mismatches (347 formulas) - investigate individually

---

## Implementation Approach

### Option 1: Post-Process Formula Rewriting (Recommended)
```python
def fix_indirect_address_formulas(ods_path: Path, excel_path: Path):
    """Fix INDIRECT/ADDRESS formulas after native conversion."""

    # 1. Extract all formulas with INDIRECT/ADDRESS from Excel
    formulas_to_fix = find_indirect_address_formulas(excel_path)

    # 2. For each formula, use LLM to translate to Calc-compatible syntax
    for cell_addr, excel_formula in formulas_to_fix:
        calc_formula = llm_translate_formula(
            excel_formula,
            context="Convert INDIRECT/ADDRESS(..., sheet) to LibreOffice Calc syntax",
            rules=INDIRECT_ADDRESS_RULES
        )

        # 3. Update formula in ODS via UNO
        update_formula_in_ods(ods_path, cell_addr, calc_formula)

    # 4. Recalculate all formulas
    recalculate_ods(ods_path)
```

### Option 2: Pre-Process Excel Formulas
- Extract formulas before native conversion
- Translate incompatible patterns
- Create modified Excel with fixed formulas
- Then run native conversion

---

## Testing

After fixes are applied, expected results:
- **Target:** ≥99% formula equivalence
- **Test:** Run `compare_excel_calc()` on fixed ODS
- **Validation:** Check that `#NAME?` errors are eliminated

---

## Files to Create

1. `src/xlsliberator/fix_formulas.py` - Formula repair module
2. `rules/formula_incompatibilities.yaml` - Pattern matching rules
3. `tests/test_formula_fixes.py` - Unit tests for formula translation
4. Integration into `api.py` as Step 1.6 (after native conversion)

---

## Example Translation Rules

```yaml
# rules/formula_incompatibilities.yaml

indirect_address_with_sheet:
  pattern: 'INDIRECT\(ADDRESS\(([^,]+),([^,]+),([^,]+),([^,]+),"([^"]+)"\)\)'
  excel_example: 'INDIRECT(ADDRESS(48+COLUMN(),12,1,1,"Spielplan"))'
  calc_replacement: 'INDIRECT("$5." & ADDRESS($1,$2,$3,$4))'
  calc_example: 'INDIRECT("$Spielplan." & ADDRESS(48+COLUMN(),12,1,1))'
  description: "ADDRESS with sheet name parameter not supported in Calc"

address_with_sheet:
  pattern: 'ADDRESS\(([^,]+),([^,]+),([^,]+),([^,]+),"([^"]+)"\)'
  excel_example: 'ADDRESS(10,5,1,1,"Sheet1")'
  calc_replacement: '"$5." & ADDRESS($1,$2,$3,$4)'
  calc_example: '"$Sheet1." & ADDRESS(10,5,1,1)'
  description: "Sheet name parameter in ADDRESS not supported"
```

---

## LLM Prompt Template

```
You are translating Excel formulas to LibreOffice Calc syntax.

Excel formula:
{excel_formula}

Issue: Excel's ADDRESS() function accepts a sheet name as the 5th parameter:
ADDRESS(row, col, abs_type, a1, "SheetName")

But LibreOffice Calc does NOT support this. Instead, use:
"$SheetName." & ADDRESS(row, col, abs_type, a1)

Also, INDIRECT() needs the full sheet-qualified address:
INDIRECT("$SheetName." & ADDRESS(...))

Translate the formula to LibreOffice Calc compatible syntax.
Preserve all logic, references, and calculations.

Output only the translated formula, no explanation.
```

---

## Estimated Impact

- **Current:** 88.21% equivalence (20,908 matching)
- **After INDIRECT/ADDRESS fix:** ~99.8% equivalence (23,660 matching)
- **After all fixes:** ~99.9% equivalence (23,679 matching)

Total repair workload: ~1,200 unique formulas to translate
