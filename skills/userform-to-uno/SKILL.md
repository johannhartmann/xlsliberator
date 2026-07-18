---
name: userform-to-uno
description: Use this skill when VBA UserForms must become native UNO dialogs, sidebars, or a capability-scoped local web interface.
compatibility: Docker-only XLSLiberator; UI behavior is exercised in LibreOffice 26.2.4.2 through real target dispatch.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# UserForm to UNO UI

Preserve user-visible workflow and accessibility, not MSForms implementation details.

## Use when

Use when the dossier contains UserForms, form controls, validation, focus order,
default/cancel buttons, or form event code. Do not use for sheet-only controls,
headless formula repair, or retaining MSForms/VBA.

## Inputs and outputs

Inputs: complete form definitions/code, images, labels, control properties,
tab/focus order, event handlers, validation/error behavior, requirements, and
screenshots when available. Outputs: UNO dialog/sidebar or local web UI,
listener/controller code, resource files, UI state contract, behavioral tests,
and visual/runtime evidence.

## Tool sequence

1. Inventory controls, properties, layout, focus/tab order, defaults, validation,
   events, modal/modeless behavior, and document state dependencies.
2. Choose UNO dialog, sidebar, or local web UI based on lifecycle,
   accessibility, complexity, and deployment; document the decision.
3. Separate presentation, validation, state, and open service calls.
4. Implement native listeners and explicit error/focus transitions.
5. Bind UI lifecycle to document events without Basic.
6. Test keyboard navigation, default/cancel buttons, invalid input, repeated
   open/close, real dispatch, dependency errors, and persistence.
7. Capture stable-region visual evidence and save/reopen results.

## Failure handling

If automation cannot trigger a required UI path, mark it `UNAVAILABLE`; calling
the handler directly is only a diagnostic. Preserve exact user-facing error
messages when required or document approved changes.

## Acceptance checklist

- [ ] All required controls and lifecycle events have target-native equivalents.
- [ ] Validation, focus, defaults, errors, and accessibility are tested.
- [ ] Real UI dispatch—not only direct handlers—drives acceptance.
- [ ] Document state persists correctly across close/reopen.
- [ ] No UserForm, MSForms, VBA, COM, or Excel facade remains.

## Tested examples

Positive: an invoice form rejects an empty customer, focuses the field, uses
Enter for Save and Escape for Cancel, then persists a valid invoice after reopen.

Adversarial: a test invokes `on_save()` directly while the Save button has no
listener. Treat the UI migration as failed.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
