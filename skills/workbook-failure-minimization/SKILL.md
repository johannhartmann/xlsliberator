---
name: workbook-failure-minimization
description: Use this skill when a reproducible migration or LibreOffice failure must be reduced to the smallest redistributable workbook regression that still fails.
compatibility: Docker-only XLSLiberator; reductions run on disposable copies with pinned LibreOffice 26.2.4.2.
recommended-tools: read_file write_file xlsprobe odstool migration-check libreoffice-runtime-mcp
---

# Workbook failure minimization

Reduce a failing case while preserving the exact failure predicate and security
boundary.

## Use when

Use after a failure is deterministic enough to test and before a generic
XLSLiberator/LibreOffice repair is promoted. Do not minimize the only source
copy, change the expected result, weaken the validator, or publish customer data.

## Inputs and outputs

Inputs: immutable source/candidate, exact failing scenario, logs, timeout and
resource limits, provenance constraints. Outputs:

- `tests/corpus/regressions/<id>/` or private equivalent;
- minimized source and dependency bundle;
- machine-readable failure predicate and reproduction command;
- reduction history, hashes, classification, and redaction/provenance record.

## Tool sequence

1. Reproduce the failure twice in a clean pinned container.
2. Define a stable predicate based on behavior/error identity, not “command was nonzero.”
3. Delta-debug job copies across sheets, ranges, modules, procedures, controls,
   formulas, names, styles, relationships, package parts, and dependencies.
4. After every reduction, validate package structure and rerun the exact predicate.
5. Preserve coupled elements when removal changes the failure class.
6. Remove or synthesize sensitive data while rechecking the predicate.
7. Run the minimized case against stock and candidate repairs and record hashes.

## Failure handling

If the failure is flaky, first isolate timing/state and report `UNRESOLVED`.
Timeout reductions must prove process-tree cleanup. If redaction destroys the
failure, keep the fixture private and publish only safe metadata. Never disable
archive, XML, or runtime limits to make minimization proceed.

## Acceptance checklist

- [ ] The original and minimized cases fail with the same predicate.
- [ ] The minimized artifact is structurally valid and reproducible.
- [ ] Reduction history explains each retained component.
- [ ] Sensitive/private data and licensing are handled.
- [ ] A failing-before/passing-after test can consume the fixture.

## Examples

Positive: reduce a 20-sheet XLSX to one sheet, one shared formula, and one style
record while preserving the same LibreOffice import assertion failure.

Adversarial: delete the assertion that distinguishes wrong output and keep only
the process exit code. Reject this as validator weakening, not minimization.
