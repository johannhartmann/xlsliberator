---
name: migration-mutation-testing
description: Use this skill when migration tests must prove they detect plausible mistranslations in generated Python, formulas, controls, or event wiring.
compatibility: Docker-only XLSLiberator; mutations run on disposable job copies and execute with LibreOffice 26.2.4.2.
recommended-tools: read_file write_file migration-check libreoffice-runtime-mcp
---

# Migration mutation testing

Measure whether public tests kill realistic migration mistakes. Never mutate the
accepted candidate or weaken acceptance to improve a score.

## Use when

Use after source-derived acceptance tests and a runnable candidate exist, before
independent review, or when a missed defect suggests a weak suite. Do not use as
a substitute for source review, hidden tests, or real target execution.

## Inputs and outputs

Inputs: immutable candidate snapshot, plan, dossier, source-derived cases, and
mutation budget. Outputs: `migration/evidence/mutations/report.json`, per-mutant
diffs/logs, killed/survived/invalid classifications, and new regression cases
for meaningful survivors.

## Tool sequence

1. Copy the candidate into a disposable mutation workspace.
2. Generate one change per mutant, prioritizing source risks:
   comparison flips; boundary shifts; wrong sheet/range; deleted state update;
   reordered event; missing listener; stale recalculation; wrong formula
   separator/reference; swallowed error; disabled save persistence.
3. Validate that each mutant is syntactically/package valid.
4. Execute the unchanged public suite in LibreOffice, including save/reopen.
5. Restore the clean snapshot between mutants and clean the process tree.
6. Classify `KILLED`, `SURVIVED`, `INVALID`, `TIMED_OUT`, or `UNAVAILABLE`.
7. For likely mistranslations that survive, add a source-derived case and rerun
   clean plus mutant candidates.

## Failure handling

Invalid mutants do not improve the kill rate. Timeouts are failures unless the
mutant intentionally models a detected loop and the sandbox proves cleanup.
Budget exhaustion leaves remaining mutations `NOT_RUN`; it never becomes a
pass.

## Acceptance checklist

- [ ] Mutants are isolated, reproducible, and recorded as diffs.
- [ ] High-risk source branches and event paths have representative mutants.
- [ ] Likely mistranslations are killed by unchanged tests.
- [ ] Survivors create source-derived test improvements or explicit unresolved findings.
- [ ] Invalid, timed-out, unavailable, and not-run counts are separate.
- [ ] Clean candidate still passes after test additions.

## Examples

Positive: mutate `quantity > 100` to `quantity >= 100`; the boundary case at
100 fails, killing the mutant.

Adversarial: replace the entire Python module with invalid syntax and count the
import failure as proof of behavioral coverage. Classify it `INVALID`; create a
plausible semantic mutant instead.
