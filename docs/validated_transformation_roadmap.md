# Validated Transformation Roadmap

XLSLiberator is moving from a best-effort Excel-to-ODS converter toward a validated transformation system for XLS, XLSM, XLSX, and XLSB workbooks:

```text
parse -> normalize IR -> transform -> emit ODS -> evaluate in LibreOffice and Apache OpenOffice -> diff -> repair -> certify
```

Validation code must never change a user's global LibreOffice or Apache OpenOffice macro security settings. Runtime checks must use isolated temporary user profiles and report skipped checks when an isolated runtime is not available.

## Staged Architecture

1. Full parse inventory
   - Inspect each source workbook and record parsed, unsupported, and failed artifacts before conversion.
   - Treat legacy `.xls` as incomplete until BIFF parsing can enumerate cells, formulas, controls, and macros honestly.

2. Rich workbook artifact IR
   - Extend the current workbook IR with formulas, controls, event bindings, unsupported artifacts, source references, target references, and validation gates.
   - Keep unsupported artifacts visible in reports instead of silently dropping them.

3. CalcBackend abstraction for LibreOffice and Apache OpenOffice
   - Discover installed office backends without failing when binaries are absent.
   - Run target validation through isolated user profiles with explicit backend metadata.

4. Formula parser/compiler/evaluator interface
   - Separate structural formula checks from target-backed FormulaParser validation.
   - Keep deterministic formula repairs in a rule registry and report unresolved mismatches.

5. Control/event inventory and GUI validation graph
   - Parse ODS XML for forms, controls, and event listeners.
   - Drive GUI validation from discovered controls and events rather than hardcoded button names.

6. Contract-driven VBA translation with a compatibility runtime
   - Preserve source-map IDs from VBA modules and procedures into generated Python.
   - Provide a small Excel compatibility runtime that translated macros can target and that can later be backed by UNO.

7. Validation gates and certification report
   - Combine inventory, formula, macro, control, backend, and runtime checks into gate results.
   - Emit machine-readable JSON and human-readable Markdown certification reports.

8. Agentic repair loop
   - Use validation evidence to propose deterministic repairs first.
   - Keep agentic repair optional, auditable, and bounded by repair history in the certification report.

## Current Implementation Gaps

- `api.py`: orchestrates conversion and validation in one path and currently needs safe isolated macro runtime handling.
- `extract_excel.py`: handles XLSX, XLSM, and limited XLSB extraction, but `.xls` parsing is currently incomplete and must not be reported as full parse success.
- `formula_ast_transformer.py`: contains a narrow Calc grammar and one deterministic INDIRECT/ADDRESS repair; it needs a rule registry and target parser seam.
- `testing_lo.py`: compares cached Excel values against Calc values but is tied to LibreOffice runtime availability.
- `python_macro_manager.py`: validates embedded Python macros and runs macro execution tests; runtime execution needs isolated profiles.
- `agent_validator.py`: performs shallow GUI validation and relies on hardcoded button names instead of discovered controls/events.
- `mcp_tools.py`: exposes current conversion and runtime helpers but needs inventory, validation, controls, and event-binding tools.
- `embed_macros.py`: rewrites Basic event URLs heuristically; it needs source-map-aware event binding support.
- `report.py`: contains `ConversionReport`; certification output should be added without breaking this report.

## Milestones

1. Add validation and artifact IR models.
2. Add safe Calc backend discovery and isolated runtime profiles.
3. Remove default global macro security mutation from conversion.
4. Add workbook inspection API, CLI, and MCP tool.
5. Add certification report writer.
6. Add formula validation seam and target parser integration.
7. Add ODS control/event inventory and source-map-aware event writing.
8. Add workbook snapshots, diffs, and validation gate runner.
9. Expose validation through CLI and MCP.
10. Replace empty `.xls` placeholder with honest legacy XLS inventory.
11. Add VBA compatibility runtime and source-map preservation.
12. Add formula rule registry and deterministic repair loop.
13. Add `transform_validated()` high-level API.
14. Refactor agent validation to consume inventories, gates, and snapshots.
