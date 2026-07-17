# XLSLiberator Implementation Status

Last updated: 2026-07-17 on `feat/evidence-certification-system`.

## Current architecture

Docker is the only application, development, test, and office-runtime platform.
The host shell is limited to Docker, Git, and file operations. LibreOffice is the
sole target and is pinned to full build `26.2.4.2`; its bundled Python, UNO, and
PyUNO execute only in disposable, resource-limited office containers.

Application entry points now require an explicitly marked Docker container.
Direct host execution fails before conversion or office-runtime construction.
The legacy direct-UNO API is a fail-closed compatibility shim. Runtime requests
resolve the pinned image to an immutable ID, verify full provenance, stage only
whitelisted workspace files, and never fall back to a host executable.

## Latest Docker-only verification

| Check | Result |
|---|---|
| Docker quality gate | 404 passed, 6 explicitly skipped, 15 integration tests deselected |
| Ruff format/lint | passed; 170 files |
| mypy | passed; 88 source files |
| pinned LibreOffice runtime probe | passed; `26.2.4.2`, Python `3.12.13`, matching PyUNO |
| real LibreOffice integration suite | 10 passed, zero skips |
| source differential | stock failed as expected; patched runtime passed |
| security gate | pip-audit: no known vulnerabilities; Bandit: zero issues |
| conformance corpus | 12 fixtures: 10 passed, 2 non-blocking unavailable, zero failed |
| release gate | ready; all 7 blocking gates passed |
| Docker/web end-to-end smoke | passed; readiness plus HTTP XLSX-to-ODS conversion |
| unauthorized application-container invocation | rejected before runtime construction |

The real integration run includes generated XLSX-to-ODS conversion, reopen and
recalculation, document inspection and repair, FormulaParser round-trip,
transactional repair, concurrent isolated scenarios, denied macro/control
capabilities, and the typed VBA compatibility backend.

## Boundary defects closed

- Removed documentation and prompt instructions that allowed host Python, `uv`,
  host LibreOffice diagnostics, or direct `soffice` fallback.
- Added a fail-closed application-container marker and startup check to the CLI,
  conversion API, MCP server, and web application.
- Kept PyUNO imports confined to the office worker and retained the startup import
  guard for non-office containers.
- Moved orchestrated Pytest temporary workbooks into the exact shared workspace;
  container `/tmp` is no longer mistaken for a host-visible bind path.
- Moved validation outputs from the read-only input mount to `/job`.
- Replaced unreliable UNO bulk range writes with dimension-checked, typed cell
  writes in both scenario and VBA compatibility paths.

## Prompt ledger and primary evidence

| Prompt | Implemented result | Primary evidence |
|---|---|---|
| 00 | Baseline, ledger, and claim audit | this ledger; generated capability and readiness reports |
| 01 | Fail-closed conversion/certification statuses | `validation_models.py`, `certification_report.py`, P0 tests |
| 02 | LibreOffice-only pinned target | `docker_runtime.py`, image lock/probe, office integration attestation |
| 03 | Blocking Docker CI | `tools/ci_check.py`, CI workflows, JUnit/attestation artifacts |
| 04 | Transactional macro refinement and embedding | refinement/embedding modules and transactional unit tests |
| 05 | Truthful MCP/GUI capabilities | `mcp_tools.py`, unavailable GUI operations, tool tests |
| 06 | Translation provenance and fail-closed cache | `translation_service.py`, translation tests/evidence |
| 07 | Scenario DSL and evidence bundle | `scenarios/`, schema tests, example scenarios |
| 08 | Explicit external Excel oracle | oracle/Windows worker modules and `docs/excel_oracle.md` |
| 09 | Docker-backed LibreOffice scenarios | `libreoffice_scenario_runner.py`, real office tests |
| 10 | Canonical artifact inventory/loss accounting | `artifact_inventory.py`, inventory/release-gate tests |
| 11 | Formula semantics and repair evidence | formula semantic/certification modules and parser integration |
| 12 | Typed VBA IR and deterministic runtime | VBA IR/parser/execution modules and conformance tests |
| 13 | Bounded agent repair | `agent_repair.py`, repair policy/evidence tests |
| 14 | Untrusted execution sandbox | `execution_sandbox.py`, malicious fixtures, security tests/model |
| 15 | Reproducible office source patch | source identities plus stock-fails/patched-passes evidence |
| 16 | Corpus, fuzzing, and minimization | corpus manifest/executions/statistics and nightly workflow |
| 17 | Capability matrix/release gate | generated matrix, release inputs, release readiness report |

Concrete generated and immutable evidence locations:

- `artifacts/ci/*-attestation.json` and JUnit XML from Docker-only CI-equivalent runs;
- `office/libreoffice/conformance/evidence/tdf-172479.json` for the source differential;
- `artifacts/office-build/*/identity.json` for source, patch, binary, Python, and PyUNO identity;
- `artifacts/certification/` for sample bundles and corpus executions;
- `docs/corpus_statistics.json`, `docs/capability_matrix.{json,md}`, and
  `docs/release_readiness.md` for generated public claims and blockers.

## Remaining work

- Release readiness is **YES**, as generated in `docs/release_readiness.md`; no
  blocking corpus or release gate remains.
- The non-blocking legacy XLS formula and XLSB controls recipes remain
  `unavailable` because their source-format materializers are not part of the
  Docker target runtime.
- Direct headless form-control dispatch and LibreOffice form serialization are
  not certified. The persisted event URI is inventoried and its exact embedded
  Python handler is executed with the live UNO document context in the isolated
  office container.
- Excel source-differential execution remains unavailable without an explicitly
  provisioned Windows Excel oracle; cached OOXML values never substitute for it.
- The optional live model gate remains unavailable without its explicit
  credential and is never counted as passed.
