# XLSLiberator implementation status

Last updated: 2026-07-18.

This file is retained as a stable link for older documentation. The active,
auditable implementation ledger is
[`agentic_implementation_status.md`](agentic_implementation_status.md).

## Current boundary

- Docker is the only application, development, test, and office-runtime
  platform.
- The host shell may perform Docker, Git, and file operations only.
- LibreOffice is the sole target and is pinned to full build `26.2.4.2`.
- LibreOffice, its bundled Python, PyUNO, and UNO run only in the pinned office
  image.
- There is no Microsoft Excel execution service, Windows worker, source-runtime
  oracle, host executable discovery, or direct `soffice` fallback.
- The deterministic core does not require a model provider or model credential.
- Open-SWE is the only agent and orchestrator.
- Pinned upstream Open-SWE and the XLSLiberator graph/API are built into the
  repository's `xlsliberator-open-swe` service.
- There is no embedded legacy agent, second repository, or repository-owned
  deterministic migration orchestrator.

## Evidence policy

Generated reports from earlier certification work are historical artifacts, not
proof that the current autonomous-migration prompt pack is complete. A capability
is considered implemented only when the active ledger links its code, tests,
exact Docker command, exit status, and any required target-runtime evidence.
Missing, skipped, unavailable, or transport-only results never count as passed.

See the active ledger for the prompt-by-prompt status and remaining blockers.
