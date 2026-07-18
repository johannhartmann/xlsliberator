---
name: vba-to-python-uno
description: Use this skill when a complete VBA project must be liberated into direct Python and UNO behavior for LibreOffice.
compatibility: Docker-only XLSLiberator; generated Python executes through the pinned LibreOffice 26.2.4.2 runtime without VBA or Excel.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# VBA to Python/UNO

Translate observable source behavior, not the VBA syntax surface.

## Use when

Use when the dossier confirms VBA modules, class modules, workbook/worksheet
events, UserForms, or shared project state. Do not use for formula-only work,
package-only edits, or any plan that retains VBA execution.

## Inputs and outputs

Read the complete VBA project, forms, event bindings, dossier, dependencies,
requirements, and acceptance cases. Never translate an isolated procedure
without resolving its project-wide references. Produce direct Python/UNO
modules, state ownership documentation, event wiring, open service adapters,
source-derived tests, and LibreOffice execution evidence.

## Tool sequence

1. Inventory modules, declarations, globals, classes, public entry points,
   callbacks, events, form code, error paths, late binding, and external APIs.
2. Build a behavior/call/state map with source locations; use it as analysis,
   not a new runtime IR.
3. Assign persistent, document, session, and invocation state explicitly.
4. Implement direct UNO services and document APIs. Extract reusable open
   services behind capability-scoped interfaces.
5. Migrate workbook/sheet/control events and UserForms to native listeners and UI.
6. Add tests from source branches, values, labels, comments, events, and user
   requirements, independent of generated function names.
7. Run import, real interactions, errors, repetition, save/close/reopen, and assertions.
8. Prove the output contains no VBA project or Basic event binding.

## Failure handling

Unsupported source behavior remains `UNRESOLVED` with source and evidence paths.
Do not add an Excel object-model facade, `ExcelContext`, a VBA interpreter, a
Windows worker, or direct provider code to make translation easier. Escalate
LibreOffice defects with a minimized fixture.

## Acceptance checklist

- [ ] Complete project and cross-module state were analyzed.
- [ ] Direct Python/UNO implements each required behavior.
- [ ] Events and UI dispatch through target-native listeners.
- [ ] Source-derived tests cover branches, errors, repetition, and persistence.
- [ ] Real LibreOffice execution and save/reopen pass.
- [ ] No VBA, Excel, COM, Windows, or facade runtime remains.

## Tested examples

Positive: `Workbook_Open` initializes shared state used by a sheet-change event
and UserForm. Implement document-scoped Python state, register listeners, drive
open/edit/form interactions, and assert persistence after reopen.

Adversarial: translate `Range("A1").Value` by growing a generic
`Excel.Application` compatibility object. Reject the design and use direct UNO
cell access in the workbook-specific behavior.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
