---
name: formula-migration
description: Use this skill when Excel formulas must be inspected, translated, recalculated, and behaviorally verified as native LibreOffice formulas.
compatibility: Docker-only XLSLiberator; formula parsing and recalculation use LibreOffice 26.2.4.2 FormulaParser in the pinned office image.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# Formula migration

Work directly with source and target formulas. A custom formula IR is not a
required architecture.

## Use when

Use for formula syntax, names, arrays, references, recalculation, errors, or
LibreOffice compatibility. Do not use for VBA translation or to replace formulas
with opaque precomputed values.

## Inputs and outputs

Inputs: all source formulas and names with locations, cached/sample values,
locale, dependencies, dossier, and behavioral requirements. Outputs: native
target formulas/names, translation rationale for exceptions, source-derived
precedent matrices, regression cases, recalc/save/reopen evidence, and classified
LibreOffice defects.

## Tool sequence

1. Read source and imported target formulas directly, preserving addresses,
   shared/array/spill structure, names, and locale.
2. Use LibreOffice FormulaParser and runtime parsing; do not guess separators or
   function mappings from strings alone.
3. Force hard recalculation and capture formula text, types, values, and errors.
4. Vary precedent values across normal, boundary, blank, text, error, date, and
   repeated-recalc cases.
5. Repair syntax or behavior with native LibreOffice formulas and reusable rules.
6. Run save/close/reopen and repeat the matrix.
7. If LibreOffice itself is wrong, minimize the workbook, add a failing
   regression, and route an upstreamable patch rather than papering over it.

## Failure handling

Separate parser, import, recalc, serialization, and test-oracle failures. Never
declare success from cached values or a single example. Keep unsupported
functions explicit and seek an open add-in/service only when native repair is
not appropriate.

## Acceptance checklist

- [ ] Source and target formulas/names are linked by address and evidence.
- [ ] FormulaParser and hard recalculation were used.
- [ ] Precedent variation covers branches and boundary behavior.
- [ ] Errors and repeated/save-reopen behavior match requirements.
- [ ] Generic fixes have regression cases.
- [ ] No custom semantic runtime or unexecuted success claim exists.

## Tested examples

Positive: vary dates around month-end for an Excel `EOMONTH` expression, parse
the target formula, force recalc, and verify values/errors before and after reopen.

Adversarial: copy cached XLSX results into ODS and remove formulas. Reject it
even if the initial visible values match.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
