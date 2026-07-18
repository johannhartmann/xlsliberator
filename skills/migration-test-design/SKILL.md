---
name: migration-test-design
description: Use this skill when public acceptance scenarios must be derived from workbook source behavior rather than the generated migration implementation.
compatibility: Docker-only XLSLiberator; migration-check scenarios execute against LibreOffice 26.2.4.2 in the pinned office container.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# Migration test design

Design independent, source-derived behavioral tests before or alongside
implementation.

## Use when

Use after forensics identifies behavior, when a plan needs executable
acceptance, or when tests fail to expose an omission. Do not use to mirror
generated function names, assert file existence only, or certify unexecuted
behavior.

## Inputs and outputs

Read original VBA, formulas, events, labels, comments, existing values, controls,
dependency contracts, dossier findings, and user requirements. Produce:

- versioned `migration/acceptance/*.yaml`;
- reusable input fixtures and explicit capability mocks;
- an evidence matrix linking each case to source locations;
- independent oracles and known observability gaps.

## Tool sequence

1. Enumerate behaviors and branches directly from source and requirements.
2. Partition normal, boundary, error, recovery, and security cases.
3. Include repeated execution, idempotency, recalculation, event order,
   interaction dispatch, save/close/reopen, and dependency failure.
4. Vary precedents and user inputs so formulas and conditions cannot pass on a
   single memorized sample.
5. Define setup, real operation, assertions, cleanup, and expected evidence for
   each scenario.
6. Run cases against a known failing or incomplete candidate when available;
   confirm they can expose missing behavior.
7. Keep public tests separate from reviewer-only hidden tests.

## Failure handling

If behavior is ambiguous, preserve competing interpretations and request a user
decision. If the runtime cannot perform a required interaction, record
`UNAVAILABLE`; do not replace it with a direct handler call. If an oracle depends
on proprietary infrastructure, define an explicit open mock plus an integration
contract.

## Acceptance checklist

- [ ] Cases cite original source evidence and user requirements.
- [ ] Branches, boundaries, errors, repetition, and save/reopen are covered.
- [ ] UI tests use real dispatch and events.
- [ ] Formula cases vary precedents and force recalculation.
- [ ] Assertions check behavior, not only output existence or generated internals.
- [ ] Skipped/unavailable required operations remain blocking.

## Examples

Positive: VBA applies a surcharge when quantity is greater than 100. Test 99,
100, 101, a blank, invalid text, two consecutive recalculations, and saved/reopened 101.

Adversarial: generated code exposes `calculate_surcharge()`. A test calling that
function directly is insufficient because the workbook event may never dispatch.
Drive the source-observable edit/event path instead.
