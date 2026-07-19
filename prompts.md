# XLSLiberator implementation contract

This contract supersedes the obsolete local-conversion web prompts.

## Non-negotiable architecture

- Docker is the only development, test, application, and LibreOffice platform.
- The host may run Docker, Git, and file operations only.
- Never run host Python, `uv`, PyUNO, UNO, LibreOffice, or `soffice`, including
  diagnostics.
- LibreOffice is the only office target and is pinned to full build `26.2.4.2`.
- Open-SWE is the only agent and only agentic orchestrator.
- There is one XLSLiberator repository.
- Do not create or require a separate Open-SWE fork.
- Embed the XLSLiberator-specific graph/API on a pinned upstream Open-SWE
  runtime in this repository's Docker stack.
- Do not add an alternate agent, provider client to the deterministic core,
  local model loop, or deterministic migration orchestrator.
- Do not use GitHub Models automatically. Provider use is an explicit Open-SWE
  operator choice and may not gate deterministic XLSLiberator operations.

## Deterministic core

XLSLiberator owns deterministic workbook inspection, raw VBA extraction, native
conversion, ODS package operations, target-native module upsert, pinned
LibreOffice execution, and evidence gates. These tools never select a model or
load a provider credential.

Source VBA without Open-SWE-produced target-native Python/UNO modules remains an
explicit unresolved capability. Native conversion, file creation, syntax, or a
successful transport response never prove behavioral equivalence.

## Web application

The web service is an authenticated client of the repository's internal
Open-SWE Compose service:

1. validate and store the upload under a server-generated owner-scoped job;
2. create or resume an Open-SWE workbook migration thread;
3. poll or stream sanitized Open-SWE stage events;
4. accept follow-up requirements and dependency uploads on the same thread;
5. propagate cancellation;
6. download owner-checked deliverables;
7. delete private local inputs according to policy.

The web container must not:

- import or call `xlsliberator.api.convert`;
- start LibreOffice, UNO, PyUNO, `soffice`, or Python subprocesses;
- receive the Docker socket;
- fall back to local conversion when Open-SWE is absent or fails;
- claim independent review unless Open-SWE supplies its evidence.

Configuration uses only the `XLSLIBERATOR_OPEN_SWE_*` namespace. Readiness
reports Open-SWE configuration and reachability separately from local storage.

## Required enforcement

Architecture tests must fail if any of the following returns:

- `src/xlsliberator/legacy_agent/`;
- `src/xlsliberator/orchestrator/`;
- `src/xlsliberator/web/orchestrator.py`;
- an `xlsliberator-orchestrator` Compose service;
- `XLSLIBERATOR_ORCHESTRATOR_*` configuration;
- a local web conversion fallback;
- a mandatory model/provider SDK dependency.

All verification commands run through Docker Compose.
