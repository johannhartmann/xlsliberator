---
name: libreoffice-debugging
description: Use this skill when import, recalculation, interaction, serialization, or runtime evidence must isolate a defect across generated output, XLSLiberator, and LibreOffice.
compatibility: Docker-only XLSLiberator; diagnostics use pinned LibreOffice 26.2.4.2 containers and never host UNO or soffice.
recommended-tools: read_file xlsprobe odstool migration-check libreoffice-runtime-mcp
---

# LibreOffice debugging

Localize the failure before changing code.

## Use when

Use after a deterministic target execution failure, timeout, crash, wrong value,
lost binding, or save/reopen regression. Do not use to start host LibreOffice,
guess from a stack trace alone, or weaken acceptance.

## Inputs and outputs

Inputs: exact source/candidate hashes, scenario, runtime/session identity,
import/runtime logs, package diff, and expected behavior. Outputs: reproducible
failure, stage localization, ownership classification, minimized next step,
bounded diagnostics, and repair recommendation.

## Tool sequence

1. Reproduce in a fresh pinned runtime with exact scenario and limits.
2. Separate package/import, load, formula parse, recalculation, listener
   registration, interaction, service call, save, and reopen stages.
3. Inspect runtime logs, UNO exceptions, package invariants, formulas, bindings,
   process status, and evidence timestamps.
4. Compare a minimal known-good target, generated target, and when relevant
   stock versus patched LibreOffice.
5. Classify ownership: workbook-specific output, XLSLiberator tool, skill,
   open-service adapter, LibreOffice, or validator.
6. Minimize generic defects and create failing regression evidence before patching.

## Failure handling

Crashes/timeouts require process-tree cleanup and preserved bounded logs. If the
failure cannot be reproduced, report `UNRESOLVED` with variables tried. Do not
attribute an empty log to absence of a defect.

## Acceptance checklist

- [ ] Exact build, hashes, stage, and scenario are recorded.
- [ ] Failure reproduces in a clean Docker runtime.
- [ ] Ownership is supported by stock/known-good/minimized comparisons.
- [ ] No host Python, UNO, PyUNO, LibreOffice, or soffice ran.
- [ ] Repair is aimed at the correct layer with a regression.

## Tested examples

Positive: package verification passes, import succeeds, but a formula changes
after reopen only on stock LibreOffice. Minimize the formula/serialization case
and route it to core patching.

Adversarial: catch a UNO exception and return the cached expected value. Reject
this as fake success.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
