# Open-SWE migration architecture

Status: accepted direction, implementation in progress
Date: 2026-07-18

## Purpose

XLSLiberator is becoming the deterministic toolbelt for an autonomous,
open-source Excel-to-LibreOffice migration system. Long-running model
orchestration belongs in a thin `xlsliberator-swe` fork of Open-SWE. The target
is always LibreOffice; Microsoft Excel is neither a runtime nor a validator.

## Binding non-goals

The system must not introduce:

1. a Microsoft Excel runtime, Windows worker, Office automation oracle, or
   proprietary validation dependency;
2. a VBA interpreter, VBA compatibility layer, Excel object-model emulator, or
   permanent Excel-shaped runtime inside LibreOffice;
3. a large custom semantic language or compiler IR that replaces direct model
   understanding of VBA, formulas, Python, or UNO;
4. a provider-specific model client in the deterministic XLSLiberator package;
5. self-certification by the implementation model;
6. success based only on an output file, document open, syntax, transport
   success, or model confidence;
7. placeholder tools that claim to perform an unavailable operation;
8. weakened, skipped, deleted, xfailed, or non-blocking tests to obtain a green
   result;
9. evidence-free universal migration claims.

Small schemas for inventories, dossiers, tool calls, scenarios, traces, evidence,
repair histories, and run records are allowed. They organize operations and
proof; they do not model the source language.

## Repository split

### `johannhartmann/xlsliberator`

This repository owns deterministic, provider-neutral domain behavior:

- inspect workbook and package artifacts;
- extract source VBA and other opaque content without interpreting it as a new
  language;
- emit a versioned migration dossier;
- perform native conversion primitives in Docker;
- edit ODS packages transactionally;
- run explicit acceptance scenarios;
- operate isolated LibreOffice sessions in the pinned runtime image;
- expose curated MCP tools for those operations;
- maintain public fixtures, migration episodes, deterministic evidence, and
  target-native helper code;
- promote reusable deterministic repairs.

It does not select models, call a provider, manage a coding-agent trajectory, or
decide that its own generated migration is acceptable.

### `johannhartmann/xlsliberator-swe`

This separate repository stays a thin customization of Open-SWE and owns:

- durable LangGraph threads and runs;
- sandbox lifecycle and workbook attachment hydration;
- model/provider routing and credentials;
- the migration lead and focused specialist subagents;
- progressive Deep Agents skills;
- curated XLSLiberator/corpus/build-farm MCP clients;
- deterministic checkpoint, evidence, anti-fake-success, and
  anti-test-weakening middleware;
- independent review, hidden tests, mutation tests, and release decisions;
- web/API/Slack/Linear/GitHub triggers;
- generic-repair pull requests and LangSmith trajectory evaluation.

The fork records its upstream commit and a rebase/sync procedure. Open-SWE
application code must not be copied into the XLSLiberator package.

## Operational boundaries

| Kind | Purpose | Examples | Must not do |
|---|---|---|---|
| Domain tool | one deterministic, inspectable operation | `xlsprobe inventory`, `odstool apply`, scenario execution | choose a model, hide failure, mutate outside its declared transaction |
| Skill | progressive instructions and reusable domain knowledge | workbook forensics, Calc formulas, VBA-to-UNO, controls | receive secrets or broad tools merely because a parent has them |
| Subagent | focused reasoning with a bounded context/toolset | formula specialist, VBA liberator, UI specialist, test adversary | certify its own work or inherit every lead-agent privilege |
| MCP | remote/sessionful execution boundary | LibreOffice runtime, hidden corpus, build farm | translate transport success into operation success |
| Middleware | deterministic trajectory policy | hydration, checkpoint, evidence requirement, no-test-weakening | make semantic claims without evidence or silently repair policy violations |

Current Deep Agents behavior is important to the design: custom subagents get
skills only when explicitly configured with skills middleware. Specialist skill
and tool scopes therefore remain explicit rather than being inherited from the
lead.

## Target command and service contracts

### `xlsprobe`

`xlsprobe` is read-only. It fingerprints the source and emits a versioned
migration dossier containing artifact inventories, formulas, VBA source,
controls, charts, names, links, external dependencies, package metadata, and
known extraction gaps. Unreadable or unsupported artifacts are preserved as
explicit unresolved entries.

### `odstool`

`odstool` plans and applies narrowly scoped package changes. Every mutation has
input hashes/preconditions, a dry-run representation, member-level changes,
conflict detection, a temporary output, post-write validation, atomic commit,
and rollback. Unknown package members and unrelated scripts are preserved.

### `migration-check`

`migration-check` executes declared acceptance scenarios against a target ODS in
the pinned LibreOffice Docker runtime. It records each open, interaction,
recalculation, save, close, reopen, package, assertion, timeout, and error state.
It never infers Excel equivalence from cached Excel values or from an output
file merely opening.

### Stateful LibreOffice runtime MCP

The MCP service creates an isolated session with a disposable LibreOffice
profile, opens one staged document, performs typed operations, captures
observations, and closes/tears down the session. The MCP process may orchestrate
Docker; workbook execution stays inside the pinned office image without the
Docker socket, host filesystem, network, or ambient secrets.

### Corpus MCP

The corpus boundary serves public cases and executes access-controlled hidden
tests. The migration agent may receive a verdict and sanitized failure evidence
but not hidden expected values or test source. Missing corpus infrastructure is
`UNAVAILABLE`.

### Build-farm MCP

The build farm accepts a LibreOffice source commit, patch series, platform, and
test declaration. It returns immutable image/binary identity, logs, SBOM, and
test evidence, or a typed failed/unavailable result. It is not a placeholder
success and is introduced only when runtime evidence points to LibreOffice as
the correct repair layer.

## Trust and security boundaries

Workbook bytes, VBA/comments, cell text, names, links, external data, extracted
strings, logs, tool output, and pull-request content are untrusted data. They are
never agent instructions.

The minimum boundaries are:

- upload/trigger layer validates size, type, hashes, and archive expansion;
- sandbox receives only the job inputs and scoped credentials it requires;
- network is denied by default and allowed destinations are job-specific;
- LibreOffice runs with a disposable profile, resource/time limits, read-only
  base image, no Docker socket, and no host office discovery;
- package tools restrict paths and reject traversal, symlink, zip-bomb, and
  precondition violations;
- the lead cannot read hidden tests or approve its own output;
- reviewer identity/context is separate from implementation;
- operation status is distinct from transport status;
- `SKIPPED`, `NOT_RUN`, `UNAVAILABLE`, timeout, inconclusive, partial, and
  unimplemented are never accepted as pass;
- evidence binds source, target, tools, runtime image, scenarios, code commit,
  and artifacts by digest.

## Current model-module migration map

Every current model-dependent path has a destination:

| Current module/path | Current role | Migration destination |
|---|---|---|
| `agent_rewriter.py` | embedded analysis/design/generation/refinement | remove from core; lead and specialist trajectories in `xlsliberator-swe` |
| `pattern_detector.py` | Anthropic-based pattern/complexity analysis | deterministic inventory remains in `xlsprobe`; semantic analysis moves to the forensics/planning skills |
| `llm_formula_translator.py` | direct formula generation/repair and legacy cache | formula specialist in `xlsliberator-swe`; deterministic mappings and target parsing remain here |
| `llm_vba_translator.py` | direct VBA-to-Python generation | VBA liberation specialist; core keeps source extraction, safe embedding, and target validation |
| `vba_reference_analyzer.py` model branch | reference/pattern analysis | deterministic lexical inventory in `xlsprobe`; model analysis in the VBA/forensics skills |
| `vba_test_generator.py` | model-generated tests | test-adversary subagent; acceptance ownership remains independent |
| `vba_translation_validator.py` | model reflection over generated code | independent reviewer in `xlsliberator-swe`, supplemented by deterministic runtime gates |
| `translation_service.py` provider adapter | provider invocation, prompt, provenance cache | remove provider adapter and orchestration; retain only reusable evidence/request schemas if still necessary |
| `vba2py_uno.py` provider selection | defaults to Anthropic | replace with deterministic tool boundary invoked using agent-produced source artifacts |
| `python_macro_manager.py` self-healing path | imports `AgentRewriter` to fix execution errors | repair loop in Open-SWE; core returns typed execution failure and evidence |
| `write_ods.py` fallback | instantiates `LLMFormulaTranslator` | accept an explicit validated formula patch/plan; never call a model |
| `api.py`, CLI, web runner/settings | `use_agent` compatibility switches and embedded translation call | submit/coordinate Open-SWE jobs or execute deterministic primitives only |
| `config.py`, README, `pyproject.toml` | Anthropic key and mandatory dependency | provider config moves to `xlsliberator-swe`; core dependency and docs become provider-neutral |

Related prohibited paths are removed, not migrated:

- `windows_excel_worker.py` and `excel_oracle.py`;
- Microsoft Excel source-runtime conformance paths;
- `runtime/` Excel object-model compatibility package;
- compiler/runtime plans whose purpose is to execute VBA semantics through an
  Excel-shaped layer.

## Migration sequence

1. Make validation and CI status semantics truthful.
2. Remove provider clients, embedded orchestration, Excel worker/oracle, and
   compatibility-runtime direction from core while preserving deterministic
   extraction/conversion/evidence behavior.
3. Deliver `xlsprobe`, `odstool`, `migration-check`, and stateful runtime MCP.
4. Create the separate thin Open-SWE fork and its pinned sandbox.
5. Add triggers, skills, specialists, curated MCP clients, and deterministic
   middleware.
6. Connect the existing web workflow to durable threads.
7. Add serious episodes, generic repair promotion, hardening, independent
   evaluation, and the first complete showcase.

Each step is independently testable. No phase may keep a forbidden runtime or
provider dependency as a permanent “temporary” fallback.

## Acceptance authority

The implementation agent proposes changes. Deterministic gates execute declared
scenarios. A separate reviewer evaluates completeness and evidence. Hidden tests
and mutation tests challenge overfitting. Only their combined, digest-bound
result may accept a migration; no component may promote its own confidence to a
certification claim.
