---
name: migration-planning
description: Use this skill when dossier evidence must become a target-native LibreOffice migration plan with specialists, risks, and acceptance gates.
compatibility: Docker-only XLSLiberator; plans target LibreOffice 26.2.4.2 without Excel, VBA, COM, or Windows runtime dependencies.
recommended-tools: read_file write_file xlsprobe migration-check
---

# Migration planning

Turn the current dossier and user requirements into
`migration/plan.md`. Planning decides what must be proven, not how to imitate
Excel.

## Use when

Use after forensics is current and before production implementation, or whenever
new evidence invalidates scope. Do not use for initial extraction, direct code
translation, package mutation, or final certification.

## Inputs and outputs

Inputs: `migration/dossier.md`, immutable sources, requirements, capabilities,
budget, and known platform constraints. Outputs:

- behavior and dependency scope;
- workbook-specific work versus generic product/LibreOffice repair;
- specialist delegation with bounded deliverables;
- target-native architecture and open replacements;
- public acceptance requirements, evidence gates, risks, and unresolved items.

## Tool sequence

1. Verify dossier coverage; send gaps back to workbook forensics.
2. Map each source behavior to target behavior and an evidence-producing test.
3. Classify work as workbook-specific, XLSLiberator defect, skill defect, open
   service gap, LibreOffice defect, or validation defect.
4. Select only relevant specialists; request independent analyses for risky
   formulas, VBA, UI, dependencies, or LibreOffice behavior.
5. Replace proprietary dependencies with explicit open capabilities. Keep
   provider adapters outside core migration logic.
6. Define implementation phases, checkpoints, mutation targets, save/reopen
   checks, reviewer inputs, budgets, and escalation conditions.
7. Write the plan and obtain user clarification only for decisions that change
   observable behavior or require new authority.

## Architecture rules

Generate direct Python/UNO, LibreOffice formulas, native controls, extensions,
or open services. Never plan an Excel worker, VBA interpreter, Excel object
model facade, expanding `ExcelContext`, custom semantic language, COM
automation, or a required proprietary Office runtime.

## Failure handling

If evidence is missing, capabilities are unavailable, or acceptance is
unobservable, mark the item blocked and name the next evidence step. If the
budget cannot cover required gates, reduce scope with user approval or finish
`UNRESOLVED`; never redefine success.

## Acceptance checklist

- [ ] Every required source behavior maps to target behavior and acceptance evidence.
- [ ] Workbook-specific and reusable work are separated.
- [ ] Specialists have narrow inputs, writable paths, and output contracts.
- [ ] Proprietary dependencies have open replacements and capability grants.
- [ ] No compatibility-layer architecture is introduced.
- [ ] Reviewer, mutation, save/reopen, security, and budget gates are explicit.

## Examples

Positive: an invoice workbook uses Outlook COM. Plan a mail-service interface,
an SMTP adapter and mock, explicit `mail` grant, UI error behavior, and an
acceptance case proving the generated attachment and requested message.

Adversarial: a request proposes implementing `Application.Range` and
`ExcelContext` so translated VBA runs unchanged. Reject it; plan direct UNO
sheet operations and source-derived behavior tests.
