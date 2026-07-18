# Autonomous migration implementation status

Last updated: 2026-07-18

This ledger tracks the Open-SWE migration prompt pack generated on 2026-07-18.
It supersedes `docs/implementation_status.md` for that migration. The older
ledger describes a different certification effort and is not evidence that any
prompt below is complete.

## Audited baseline

| Item | Audited value |
|---|---|
| XLSLiberator branch | `feat/evidence-certification-system` |
| XLSLiberator commit | `ca94903edc8f8108162c29ea28c7611354fbb32d` |
| XLSLiberator pack baseline | `12a7ccafa39ae2f9fbf48b96dd26517ba271c8c2` |
| Open-SWE `main` | `f0897479c38f2506f03b4de38081d4770928f09d` |
| Deep Agents checkout inspected | `4ffea886` |
| Target | LibreOffice only, full build `26.2.4.2` |
| Execution boundary | Docker only; no host Python, PyUNO, UNO, LibreOffice, or `soffice` |

The Open-SWE commit still matched the prompt-pack baseline at audit time. The
Open-SWE audit used a disposable checkout at
`/Volumes/CrucialMusic/src/tmp/open-swe-prompt00`; it is not the future
`xlsliberator-swe` repository.

## Baseline commands and outcomes

The required baseline was attempted from the repository root. No host Python or
office executable was invoked.

| Surface | Exact command | Exit | Outcome |
|---|---|---:|---|
| Installation/build | `docker compose build test` | 1 | **UNAVAILABLE**. Docker BuildKit failed while committing a layer: `write /var/lib/docker/buildkit/containerd-overlayfs/metadata_v2.db: input/output error`. |
| Formatting | `docker compose run --rm test ruff format --check .` | not run | **BLOCKED** by the missing test image and Docker storage I/O failure. |
| Lint | `docker compose run --rm test ruff check .` | not run | **BLOCKED** by the same Docker failure. |
| Mypy | `docker compose run --rm test mypy src/` | not run | **BLOCKED** by the same Docker failure. |
| Unit tests | `docker compose run --rm test pytest -m "not integration"` | not run | **BLOCKED** by the same Docker failure. |
| LibreOffice integration | `make test-integration` | not run | **BLOCKED**. The pinned office image is not present and Docker cannot read or commit required image data. |
| Docker/web smoke | `docker compose --profile ci-orchestrator run --rm test-orchestrator python tools/ci_check.py docker-web` | not run | **BLOCKED** by the same Docker failure. |

Read-only diagnostics established that this is a Docker Desktop storage problem,
not a project test failure:

- `docker info` reached Docker Desktop 29.2.1 on `aarch64`.
- `docker system df` exited non-zero while opening a content blob with
  `input/output error`.
- neither `xlsliberator-test:py311` nor
  `xlsliberator-libreoffice:26.2.4.2` exists locally;
- the unrelated `kind-registry` container was running, so no destructive prune
  or Docker Desktop restart was attempted;
- the host data volume had about 40 GiB free and the project volume about
  777 GiB free.

This record is not a passing baseline. Until Docker storage is repaired, all
unexecuted gates remain `BLOCKED`, never `PASSED`.

## Current architecture problems

| Problem | Current evidence | Required destination |
|---|---|---|
| Model calls in the core package | direct Anthropic use in `agent_rewriter.py`, `llm_formula_translator.py`, `llm_vba_translator.py`, `pattern_detector.py`, `vba_reference_analyzer.py`, `vba_test_generator.py`, `vba_translation_validator.py`, plus the adapter in `translation_service.py` | Prompt 02 removes orchestration/provider clients from core; specialist work moves to `xlsliberator-swe` |
| Hardcoded provider/model dependencies | mandatory `anthropic` dependency; `claude-sonnet-4-5` strings; Anthropic environment/config fields | model selection and credentials belong to the Open-SWE fork |
| Invalid model output can be returned or cached | the legacy formula translator caches minimally normalized text without target parsing and returns the original formula on failure; old caches are untyped | Prompt 01 makes every invalid/unavailable outcome fail closed; Prompt 02 retires the legacy model path |
| Excel-shaped runtime | `runtime/`, `vba_execution.py`, and `lo_worker.py` implement or consume `Application`, workbook, worksheet, range, and worksheet-function compatibility objects | remove rather than extend; generated code must call target-native UNO APIs |
| Prohibited Excel worker/oracle | `windows_excel_worker.py`, `excel_oracle.py`, `vba_conformance.py`, tests, and documentation still define Microsoft Excel source-runtime evidence | remove in Prompt 02; scenarios use declared expectations, public fixtures, and independent target tests without Excel execution |
| Fail-open or ambiguous validation states | macro/control gates can report `SKIPPED` for missing output; legacy translation fallbacks return source text; some test paths accept skipped live behavior | Prompt 01 makes required gates reject `SKIPPED`, `UNAVAILABLE`, `NOT_RUN`, timeouts, and transport-only success |
| Integration coverage is incomplete outside CI | GitHub office and web jobs are now blocking with `XLSLIBERATOR_FAIL_ON_SKIP=1`, but `make all` omits office, web, and package gates; multiple integration modules still contain permissive skips outside that mode | Prompt 01 aligns local and CI truth and separates explicitly optional live-provider tests |
| Placeholder or inaccurate tools | `take_screenshot` and keyboard input are registered MCP tools but only return unavailable; “click” resolves a handler and invokes a script rather than dispatching a GUI click; `test_placeholder.py` adds no behavior coverage | Prompt 01 corrects naming/status; Prompts 04 and 06 expose only operations with truthful semantics |
| ODS mutation risk | legacy package/script mutation has historically replaced package members wholesale; current transactional code is safer but is not yet a general package editor with preconditions, dry-run plans, rollback, and conflict detection | Prompt 04 implements transactional `odstool` and preservation tests |
| Demo evidence is too basic | the only checked-in workbook scenario is `tests/fixtures/scenarios/basic.ods`; generated samples focus on isolated features rather than serious migration episodes | Prompts 19 and 23 add representative episodes, hidden checks, mutations, and review evidence |
| Self-review risk | current embedded `AgentRewriter` generates and refines its own result | Prompts 13, 16, and 17 separate specialists, lead orchestration, and an independent reviewer |
| Large custom semantic layer | `vba_ir.py`, `vba_parser.py`, execution plans, and compatibility-runtime metadata are growing into a hand-built semantic/runtime route | Prompt 02 retains only small inventories/evidence schemas and removes compiler/runtime semantics used to emulate VBA |

## Target surfaces

| Surface | Owner | Contract |
|---|---|---|
| `xlsprobe` | `xlsliberator` | read-only workbook/package forensics and a versioned migration dossier |
| `odstool` | `xlsliberator` | preconditioned, dry-runnable, transactional ODS package mutations with rollback and preservation evidence |
| `migration-check` | `xlsliberator` | execute explicit acceptance scenarios against the target and emit deterministic traces/evidence |
| LibreOffice runtime MCP | `xlsliberator` | stateful, isolated sessions in the pinned Docker office image; operations report their own status |
| corpus MCP | separate service or `xlsliberator` boundary | public cases plus access-controlled hidden tests; never disclose hidden expectations to the implementation agent |
| build-farm MCP | separate service | build and test declared LibreOffice commits/patches and return immutable artifacts or `UNAVAILABLE` |

## Prompt checklist

`PENDING` means no completion claim has been made. A prompt becomes `COMPLETE`
only when its implementation, named commands with exit status, runtime evidence
when applicable, and ledger update are linked.

| Prompt | Status | Required acceptance evidence |
|---:|---|---|
| 00 — baseline and architecture | COMPLETE WITH BLOCKED RUNTIME | this baseline; [migration architecture](architecture/open-swe-migration.md); [decision log](architecture/decision-log.md) |
| 01 — truthful validation and CI | **NEXT** | fail-closed tests, aligned Docker CI commands, exact results |
| 02 — extract model orchestration | PENDING | dependency/import audit, removed prohibited runtime/worker paths, tests |
| 03 — `xlsprobe` dossier | PENDING | CLI/API schema, fixture snapshots, dossier evidence |
| 04 — transactional `odstool` | PENDING | mutation-plan, rollback, preservation and conflict tests |
| 05 — `migration-check` | PENDING | scenario schema, target execution traces, negative cases |
| 06 — stateful LibreOffice MCP | PENDING | session lifecycle, isolation, runtime integration evidence |
| 07 — thin Open-SWE fork | PENDING | separate repository, upstream record, sync procedure |
| 08 — sandbox snapshot | PENDING | image/SBOM, tool versions, sandbox smoke |
| 09 — triggers and hydration | PENDING | attachment tests, artifact hashes, durable-thread evidence |
| 10 — Deep Agents skills | PENDING | skill discovery/loading tests and isolation proof |
| 11 — forensics/planning/testing/package skills | PENDING | skill fixtures and trajectory tests |
| 12 — migration-specialist skills | PENDING | formula/VBA/UI/dependency/LO skill tests |
| 13 — specialist subagents and routing | PENDING | routing tests, model config, isolated skill/tool scopes |
| 14 — curated MCP tools | PENDING | allowlist, typed failures, integration traces |
| 15 — deterministic middleware | PENDING | checkpoint and anti-fake-success/anti-test-weakening tests |
| 16 — migration lead | PENDING | workflow tests and end-to-end thread trajectory |
| 17 — independent reviewer | PENDING | separate reviewer identity, hidden/mutation gate evidence |
| 18 — web to Open-SWE threads | PENDING | job/thread mapping, resume, artifact delivery tests |
| 19 — serious demos and corpus | PENDING | representative episodes and corpus metadata |
| 20 — repair promotion/build farm | PENDING | generic repair PR flow and build-farm contract/tests |
| 21 — execution hardening | PENDING | threat-model tests, sandbox/network/secret controls |
| 22 — LangSmith evaluations | PENDING | datasets, evaluators, thresholds, release gate |
| 23 — autonomous showcase | PENDING | reproducible full episode, target evidence, independent verdict |

## Next action

Run **Prompt 01 — Make the current validation and CI truthful**. Docker-backed
verification must be rerun after the Docker Desktop storage fault is repaired;
no local Python or office fallback is permitted.
