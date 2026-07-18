---
name: libreoffice-core-patching
description: Use this skill when a minimized regression proves LibreOffice itself needs an upstreamable source patch and stock-versus-patched validation.
compatibility: Docker-only XLSLiberator build farm; source and runtime commits, configuration, and LibreOffice 26.2.4.2 ancestry are recorded.
recommended-tools: read_file write_file libreoffice-build-farm-mcp migration-check
---

# LibreOffice core patching

Patch LibreOffice only after ownership is proven with a minimized case.

## Use when

Use for a reproducible LibreOffice defect that cannot be correctly repaired in
the workbook, XLSLiberator, skill, service, or validator. Do not use for
workbook-specific output, speculative fixes, or validator weakening.

## Inputs and outputs

Inputs: minimized redistributable regression, failing scenario, ownership
evidence, pinned source commit/build configuration, and authorized build-farm
grant. Outputs: isolated source worktree, focused upstream test, minimal patch,
component/runtime artifacts and hashes, stock/patched comparison, logs, clean
patch series, regression workbook, and upstream-ready draft.

## Tool sequence

1. Verify authorization and create an isolated build-farm worktree at the pinned commit.
2. Reproduce stock failure and add the smallest focused upstream test.
3. Identify the owning component and implement the minimal general fix.
4. Build the affected component and, when necessary, a runtime; record commands,
   configuration, commit, toolchain, and hashes.
5. Run focused upstream tests, affected component suite, minimized workbook, and
   relevant public corpus.
6. Compare stock and patched behavior with identical scenarios and environments.
7. Export a clean patch and regression artifact separate from generated workbook output.
8. Update debugging/migration skill guidance when the repair teaches a reusable rule.

## Failure handling

Unavailable build capacity or authorization is `UNAVAILABLE`. A patch that only
special-cases the customer workbook, weakens a test, or lacks failing-before
evidence is rejected. Keep build credentials and hidden fixtures server-side.

## Acceptance checklist

- [ ] Minimized stock failure and focused upstream test exist.
- [ ] Patch changes the correct component and is generally applicable.
- [ ] Stock fails and patched build passes the identical scenario.
- [ ] Affected upstream tests and corpus pass.
- [ ] Commits, builds, hashes, logs, and patch are reproducible.

## Tested examples

Positive: a minimized ODS proves an event binding is dropped during save. Add an
upstream serialization test, patch the component, and compare stock/patched reopen.

Adversarial: change `migration-check` to ignore the missing binding. Reject it as
test weakening, not a LibreOffice repair.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
