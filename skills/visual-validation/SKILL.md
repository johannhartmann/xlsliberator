---
name: visual-validation
description: Use this skill when workbook appearance or interactive UI behavior requires screenshot evidence and stable-region review in addition to deterministic acceptance.
compatibility: Docker-only XLSLiberator; screenshots come from the pinned LibreOffice 26.2.4.2 runtime with recorded rendering configuration.
recommended-tools: read_file write_file migration-check libreoffice-runtime-mcp
---

# Visual validation

Use visual evidence to diagnose and review layout while keeping deterministic
behavioral assertions authoritative wherever possible.

## Use when

Use for dashboards, forms, controls, charts, print layouts, conditional
formatting, or other user-visible requirements. Do not use a whole-image
similarity score as the sole acceptance oracle or infer successful interaction
from a static screenshot.

## Inputs and outputs

Inputs: source visual references when legally available, target candidate,
required views/states, stable-region masks, renderer configuration, and
behavioral expectations. Outputs: privacy-safe screenshots, region manifests,
comparison metrics, multimodal findings, deterministic corroborating evidence,
and unresolved visual differences.

## Tool sequence

1. Record LibreOffice build, fonts, locale, DPI, zoom, theme, viewport, and data state.
2. Drive the real UI/event sequence to each named state.
3. Capture bounded screenshots and define stable regions; mask clocks, cursors,
   volatile data, and platform decoration with justification.
4. Compare geometry, clipping, visibility, focus, labels, state, and print areas.
5. Ask multimodal review for structured observations, not certification.
6. Corroborate critical findings with UNO properties, event traces, exported
   output, or source-derived acceptance assertions.
7. Recheck after save/close/reopen.

## Failure handling

Missing fonts, nondeterministic rendering, inaccessible interaction, or unstable
regions remain explicit. Rebaseline only when a reviewed requirement changes,
never merely because a candidate differs. Keep screenshots private if they
contain customer data.

## Acceptance checklist

- [ ] Rendering configuration and state are reproducible.
- [ ] Stable and masked regions are documented.
- [ ] Required interactions occurred through real dispatch.
- [ ] Critical visual claims have deterministic corroboration.
- [ ] Multimodal findings are diagnostic and independently reviewed.
- [ ] Save/reopen visual state is checked where required.

## Examples

Positive: capture an invoice form before and after invalid input, compare label,
focus, default-button, and error-message regions, and confirm focus/event state
through the runtime trace.

Adversarial: a screenshot resembles the source, but the button listener was
called directly by a test. Refuse acceptance until real UI dispatch produces the
state and evidence.
