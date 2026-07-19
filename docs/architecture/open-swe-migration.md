# Open-SWE migration architecture

Status: accepted
Date: 2026-07-19

## Decision

Open-SWE is XLSLiberator's only agent and only agentic orchestrator.
XLSLiberator remains a deterministic, provider-neutral toolbelt.

There is no second XLSLiberator repository, no maintained Open-SWE fork, no
repository-owned deterministic migration orchestrator, and no alternate or
legacy model agent. The only embedded agent surface is the XLSLiberator graph
loaded on pinned upstream Open-SWE. Deterministic tool sequencing inside a
command or validation gate is ordinary application logic, not an alternative
agent.

## Runtime boundary

The browser-facing Docker service:

1. validates and stores the upload under an owner-scoped job;
2. creates or resumes an authenticated Open-SWE workbook thread;
3. streams sanitized Open-SWE events;
4. downloads owner-checked delivery artifacts;
5. deletes private local inputs according to retention policy.

It never imports the conversion API, starts LibreOffice or PyUNO, receives the
Docker socket, or falls back to local conversion.

The internal `xlsliberator-open-swe` service owns:

- durable agent state and resumability;
- model and provider selection;
- specialist routing and bounded repair;
- sandbox execution;
- independent review;
- the final migration decision.

XLSLiberator owns:

- bounded workbook inspection and raw VBA extraction;
- deterministic ODS package operations;
- the pinned LibreOffice `26.2.4.2` Docker target;
- explicit acceptance scenarios and evidence;
- fail-closed certification gates.

## Source and dependency policy

This repository does not maintain a customized Open-SWE source fork. Its Docker
image downloads an exact upstream commit, verifies the archive hash, installs
upstream's locked dependency set, and then installs the XLSLiberator-specific
graph and authenticated API from this repository. There is no runtime dependency
on another XLSLiberator repository or separately configured Open-SWE service.

The web transport reaches `xlsliberator-open-swe:2024` only on the Compose
network. The service exposes versioned routes under
`/api/xlsliberator/migrations`, persists its local LangGraph thread database and
private thread workspaces on the shared Open-SWE workspace mount, and delegates
office work only to the curated MCP gateway. The agent has no shell backend and
no Docker socket.

## Provider and cost policy

No deterministic XLSLiberator operation reads a model credential. Open-SWE
provider use must be explicitly configured by the operator. GitHub Models is
not an automatic fallback, cannot be a release or execution gate, and must not
be invoked merely because GitHub authentication exists.

An empty `XLSLIBERATOR_OPEN_SWE_MODEL` keeps the stack healthy but rejects every
migration start before a LangGraph run is created. GitHub Models additionally
requires `XLSLIBERATOR_GITHUB_MODELS_ENABLED=1` and a separate
`GITHUB_MODELS_TOKEN`.

## Fail-closed invariants

- Missing explicit model configuration means the web job fails before a run.
- Open-SWE unavailability means the web job fails.
- Missing target-native VBA artifacts remains unresolved.
- Missing independent review is not approval.
- Missing LibreOffice execution evidence is not success.
- Host Python, UNO, PyUNO, LibreOffice, and `soffice` are never used.
