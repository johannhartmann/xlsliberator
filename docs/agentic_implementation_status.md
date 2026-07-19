# Open-SWE implementation status

Last updated: 2026-07-19.

This ledger records the currently supported architecture. Evidence from the
retired separate-fork and deterministic-orchestrator experiments is historical
only and is not a product dependency or proof of the current deployment.

## Accepted architecture

- Open-SWE is the only agent and only agentic orchestrator.
- XLSLiberator is a deterministic, provider-neutral toolbelt.
- There is one XLSLiberator repository.
- There is no maintained Open-SWE fork.
- There is no embedded `legacy_agent` package.
- There is no repository-owned deterministic migration orchestrator.
- The web service delegates exclusively to Open-SWE and has no local conversion
  fallback.
- Docker is the only supported application, test, and runtime platform.
- LibreOffice is the sole target and is pinned to full build `26.2.4.2`.
- Host Python, `uv`, PyUNO, UNO, LibreOffice, and `soffice` are prohibited.

## Implemented in this repository

| Surface | Current state |
|---|---|
| Deterministic primitives | workbook inspection, native conversion, VBA extraction, ODS upsert, target lifecycle, and evidence gates |
| LibreOffice boundary | pinned Docker image and disposable runtime profiles |
| Open-SWE runtime | archive-hash-verified upstream commit built into `xlsliberator-open-swe` |
| Agent graph | thread-confined Deep Agents graph with forensics, target-native migration, and independent-review specialists |
| Agent tools | curated MCP tools only; no shell backend and no Docker socket |
| Web client | authenticated owner-scoped Open-SWE migration client |
| Web safety | bounded uploads, sanitized public state, safe artifact names, cleanup, no Docker socket |
| Failure behavior | absent/unreachable Open-SWE or missing model fails closed; no local conversion |
| Architecture guards | tests reject legacy/alternate agents, a deterministic orchestrator, old env names, and local web fallback |
| Provider dependency | no model/provider SDK in package dependencies |

## Remaining acceptance work

The repository now contains and boots the versioned Open-SWE workbook API under
`/api/xlsliberator/migrations`. The Docker smoke builds the pinned upstream
runtime, loads the real graph, verifies web-to-agent readiness, and proves the
no-model/no-cost fail-closed gate without a fake Open-SWE server.

An end-to-end model-driven workbook acceptance run remains opt-in because it
incurs provider usage:

```bash
docker compose up -d --build xlsliberator-web
docker run --rm --network xlsliberator_default --env-file .env \
  -e XLSLIBERATOR_ALLOW_PAID_OPEN_SWE_TEST=1 \
  -v "$PWD:/workspace" -w /workspace xlsliberator-test:py311 \
  pytest -m "integration and live" tests/integration/test_open_swe_real.py
```

Until that test produces current evidence, the repository must not claim that
arbitrary workbook migrations are behaviorally complete.

## Provider and cost rule

Provider use is an explicit Open-SWE operator decision. GitHub Models is not an
automatic default or fallback, cannot gate deterministic execution, and must
not be invoked solely because GitHub credentials or paid usage are available.
