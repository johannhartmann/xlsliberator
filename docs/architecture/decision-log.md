# Architecture decision log

This log records decisions for the Open-SWE migration. Decisions are append-only;
later changes supersede an entry explicitly rather than rewriting history.

## ADR-001 — LibreOffice is the only executable office target

- Date: 2026-07-18
- Status: accepted

LibreOffice full build `26.2.4.2`, including its matching Python/PyUNO stack, runs
only inside the pinned repository Docker image. Host Python, PyUNO, UNO,
LibreOffice, and `soffice` are prohibited even for diagnostics.

Consequences:

- no host executable discovery or local fallback;
- runtime failure is reported as `UNAVAILABLE` or `FAILED`;
- Docker storage/runtime availability is a real prerequisite, not a reason to
  bypass the boundary.

## ADR-002 — No Microsoft Excel execution or source oracle

- Date: 2026-07-18
- Status: accepted

The product will not build or call a Windows Excel worker, Office automation
service, or proprietary source-runtime oracle.

Consequences:

- current Excel worker/oracle modules and their tests/docs are removed;
- acceptance scenarios use declared expected behavior, public corpus evidence,
  target execution, independent hidden tests, and mutations;
- cached source values may inform forensics but cannot certify behavior.

## ADR-003 — No VBA compatibility runtime or Excel object model

- Date: 2026-07-18
- Status: accepted

Generated solutions are target-native Python/UNO, Calc/OpenFormula, UNO
dialogs/controls, LibreOffice extensions, or open replacement services. A VBA
interpreter, compatibility layer, or Excel-shaped object model is not a product
layer.

Consequences:

- the current `runtime/` direction and dependent execution plans are retired;
- agents read the original source and write native target code;
- small evidence schemas remain allowed, but no replacement semantic language
  is introduced.

## ADR-004 — Split deterministic tools from agent orchestration

- Date: 2026-07-18
- Status: accepted

`xlsliberator` owns deterministic workbook tools and LibreOffice execution.
`xlsliberator-swe`, a separate thin Open-SWE fork, owns models, trajectories,
specialists, middleware, review, and integrations.

Consequences:

- no provider SDK is mandatory in XLSLiberator core;
- Open-SWE code is not copied into this package;
- the fork pins/records upstream and maintains a sync procedure;
- the interface between repositories is versioned files, commands, and MCP
  contracts rather than Python imports across repositories.

## ADR-005 — Independent evidence decides acceptance

- Date: 2026-07-18
- Status: accepted

The model or subagent that implements a migration cannot certify it.

Consequences:

- deterministic target scenarios, hidden tests, mutation tests, and a separate
  reviewer are required;
- transport success, file creation, open success, syntax, and confidence are
  insufficient;
- skipped, unavailable, not-run, partial, timed-out, and inconclusive states
  remain non-passing.

## ADR-006 — Use operational schemas, not a semantic compiler IR

- Date: 2026-07-18
- Status: accepted

Versioned dossiers, inventories, mutation plans, scenarios, traces, evidence
manifests, and repair histories are appropriate. A custom language intended to
encode VBA or formula semantics is not.

Consequences:

- schemas must map directly to observable artifacts, tool inputs, or evidence;
- foundation models receive original formulas/VBA and relevant generated code;
- abstractions that exist primarily to emulate VBA execution are removed.

## ADR-007 — Deep Agents skills and subagent privileges are explicit

- Date: 2026-07-18
- Status: accepted

The inspected Deep Agents API does not automatically give custom subagents the
parent's skills. Each specialist gets an explicit skill, tool, middleware, and
permission scope.

Consequences:

- least privilege is enforceable and testable;
- the lead delegates focused tasks rather than sharing ambient credentials and
  tools;
- specialist responses are evidence-bearing inputs, not certification.

## ADR-008 — MCP transport and operation status are separate

- Date: 2026-07-18
- Status: accepted

An MCP request completing does not imply that the named workbook operation
succeeded.

Consequences:

- every response carries typed operation status and evidence;
- unavailable screenshot, input, GUI dispatch, corpus, or build-farm behavior
  is reported accurately and cannot satisfy an acceptance gate;
- inaccurately named legacy tools are renamed, implemented, or removed.

## ADR-009 — ODS editing is transactional and preservation-first

- Date: 2026-07-18
- Status: accepted

ODS package changes use explicit plans, preconditions, temporary outputs,
post-write checks, atomic commit, and rollback.

Consequences:

- unknown members and unrelated scripts are preserved;
- no destructive wholesale script/package replacement;
- concurrent/stale plans fail on hash conflict;
- dry-run and member-level evidence are first-class outputs.

## ADR-010 — Docker baseline failure remains a blocker, not a local fallback

- Date: 2026-07-18
- Status: accepted

The Prompt 00 build failed because Docker Desktop's BuildKit/containerd overlay
metadata returned `input/output error`. No suitable cached test or office image
was present.

Consequences:

- baseline build and test gates are recorded as blocked rather than passed;
- no destructive Docker prune/restart is performed while unrelated workloads
  are active without explicit authority;
- implementation may continue where static/file work is possible, but every
  Docker-dependent claim must be rerun after the storage fault is repaired.
