---
name: activex-to-open-controls
description: Use this skill when ActiveX or Excel Form controls must be replaced with native LibreOffice controls and verified listeners.
compatibility: Docker-only XLSLiberator; controls execute through LibreOffice 26.2.4.2 without ActiveX, COM, VBA, or Windows DLLs.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# ActiveX to open controls

Replace control behavior with native LibreOffice controls/listeners and prove
real interaction dispatch.

## Use when

Use for buttons, lists, combo boxes, checkboxes, option buttons, scroll/spin
controls, and sheet-embedded control events. Do not use for VBA UserForms or
static drawing shapes without behavior.

## Inputs and outputs

Inputs: control inventory, properties, anchors, linked cells/ranges, groups,
event bindings, handler source, z-order, and requirements. Outputs: native
controls, listener modules, mapping/evidence table, source-derived interaction
tests, and proof that legacy bindings are absent.

## Tool sequence

1. Correlate package relationships, control definitions, drawings, linked cells,
   and VBA handlers; empty extraction is not proof of absence.
2. Select native control type and preserve meaningful state, grouping, anchoring,
   labels, accessibility, and keyboard behavior.
3. Implement direct UNO listeners and target-native handler logic.
4. Remove ActiveX binaries and legacy event bindings transactionally.
5. Dispatch real mouse/keyboard/control actions and capture listener/event traces.
6. Verify linked state, repeated actions, errors, save/close/reopen, and package integrity.

## Failure handling

If runtime automation cannot dispatch a control, keep acceptance blocked. Direct
handler calls may isolate logic but cannot prove wiring. Unsupported proprietary
controls require an explicit target-native redesign, not a Windows bridge.

## Acceptance checklist

- [ ] Every behavioral control maps to a native control and listener.
- [ ] Anchoring, state, groups, labels, and accessibility are preserved as required.
- [ ] Tests distinguish real dispatch from handler invocation.
- [ ] ActiveX binaries and Basic/VBA bindings are absent.
- [ ] LibreOffice interaction and save/reopen pass.

## Tested examples

Positive: replace an ActiveX checkbox linked to `Inputs.B7`; real click updates
the cell, recalculates dependent formulas, and retains state after reopen.

Adversarial: leave the ActiveX object in the package and add a Python function
that tests call directly. Reject it.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
