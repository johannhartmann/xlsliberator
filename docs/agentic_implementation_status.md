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

## Prompt 01 verification

Prompt 01 is implemented on `feat/evidence-certification-system`; its remote
Docker CI is pending. The implementation now requires both a successful
`ConversionReport` and a structurally valid ODS ZIP package, retains that report
in certification metadata, fails missing-output gates, separates transport from
operation status in web responses, and blocks downloads until the operation is
actually `PASSED`. Required office/web/package checks are part of `make all`,
and required GitHub jobs retain logs without suppressing failures.

| Exact command | Exit | Outcome |
|---|---:|---|
| `git diff --check` | 0 | Static patch integrity passed. |
| `docker compose config --quiet` | 0 | Compose configuration parsed successfully. |
| `docker compose build test` | 1 | **UNAVAILABLE**. Docker Desktop failed to write `/var/lib/desktop-containerd/daemon/io.containerd.metadata.v1.bolt/meta.db` with `input/output error`. |
| `docker image ls --format '{{.Repository}}:{{.Tag}} {{.ID}}'` | 1 | **UNAVAILABLE**. Docker Desktop failed to open a content blob with `input/output error`. |
| Docker lint, typecheck, unit, office, web, package | not run locally | **BLOCKED** by the Docker storage corruption; remote Docker CI is the required verification path. |

The new fail-closed regression tests are
`tests/unit/test_ci_truthfulness.py`, the conversion/validation tests in
`tests/unit/test_validation_runner.py` and `tests/unit/test_validated_api.py`,
and the web operation-status/package tests under `tests/unit/web/`.

## Prompt 02 verification

Prompt 02 is implemented and locally verified in Docker. Provider-driven
translation, the semantic VBA runtime, and the prohibited Microsoft Excel
worker/oracle were removed from the deterministic package surface. Historical
provider code is isolated in the deprecated, optional
`xlsliberator.legacy_agent` namespace. The public core exposes typed,
provider-neutral primitives for inspection, native conversion, raw VBA
extraction, caller-supplied script upsert, package validation, target lifecycle,
and acceptance scenarios.

The normal Compose build remained unusable because Docker Desktop had cached an
invalid architecture layer for the pinned Python base. Validation therefore
built the same repository test Dockerfile explicitly for `linux/arm64`; no host
Python, office executable, UNO, or PyUNO process was used.

| Exact command | Exit | Outcome |
|---|---:|---|
| `docker build --platform linux/arm64 --file docker/test/Dockerfile --build-arg PYTHON_BASE=python:3.11-slim --tag xlsliberator-test:local-arm64 .` | 0 | Test image built successfully. |
| `docker run ... xlsliberator-test:local-arm64 pytest -p no:cacheprovider -q tests/unit/test_core_architecture.py tests/unit/test_primitives.py tests/unit/test_agent_repair.py tests/unit/test_translation_service.py tests/unit/test_runtime.py tests/unit/test_vba_project_ir.py tests/unit/test_vba_semantic_runtime.py` | 0 | 39 focused architecture, primitive, and legacy-boundary tests passed. |
| `docker run ... xlsliberator-test:local-arm64 pytest -p no:cacheprovider -q -m 'not integration' tests/unit` | 0 | 409 passed and 5 explicitly declared live/fixture skips. |
| `docker run ... xlsliberator-test:local-arm64 mypy --cache-dir=/tmp/mypy-cache src/` | 0 | Strict typing passed for 87 source files. |
| `docker run ... xlsliberator-test:local-arm64 ruff check --no-cache src tests` | 0 | Lint passed. |
| `docker run ... xlsliberator-test:local-arm64 ruff format --no-cache --check .` | 0 | All 169 files were formatted. |

## Prompt 03 verification

Prompt 03 is implemented and locally verified in Docker. `xlsprobe` exposes all
eight required read-only commands and a provider-neutral `ProbeReport` schema.
The transactional dossier snapshots the source, preserves exact raw workbook
and extracted VBA bytes, groups formulas, inventories package parts or OLE
streams, detects target-specific dependencies, records extractor gaps, and
delimits workbook-derived data as untrusted. Nested archives are retained as
raw evidence rather than recursively expanded.

Generated XLSX, XLSM, XLSB, and XLS fixtures cover format-specific success and
partial/unavailable extraction. Adversarial tests cover unsafe package paths,
compression ratio, source size, timeout, source mutation during snapshot,
non-overwrite behavior, truthful empty VBA results, and all required dependency
categories.

| Exact command | Exit | Outcome |
|---|---:|---|
| `docker compose run --rm test pytest -p no:cacheprovider -q tests/unit/test_xlsprobe.py` | 0 | 15 focused forensics, fixture, CLI, dossier, and limit tests passed. |
| `docker compose run --rm test pytest -p no:cacheprovider -q -m "not integration" tests/unit` | 0 | 426 passed and 5 explicitly declared live/fixture skips. |
| `docker compose run --rm test mypy --cache-dir=/tmp/mypy-cache src/` | 0 | Strict typing passed for 88 source files. |
| `docker compose run --rm test ruff check --no-cache src tests` | 0 | Lint passed. |

## Prompt 04 verification

Prompt 04 is implemented and locally verified in Docker. `odstool` exposes all
nine required deterministic commands. Mutations require a verified source,
support dry runs and SHA-256 preconditions, write a complete candidate beside
the source, fsync and re-verify it, reject concurrent source changes, and only
then atomically replace the original. Failed writes, malformed packages,
unresolved bindings, and changed preconditions leave the source untouched.

The package layer preserves unrelated scripts, unknown members, ZIP metadata,
unrelated manifest entries, XML namespaces, comments, and processing
instructions. It validates script syntax and exported event targets, enforces
ODS mimetype placement and storage, reports member-level diffs, and explicitly
reports signature invalidation. The historical embed/remove API now delegates
to this single transactional implementation.

| Exact command | Exit | Outcome |
|---|---:|---|
| `docker compose run --rm test pytest -p no:cacheprovider -q tests/unit/test_odstool.py tests/unit/test_embed_macros_transactional.py tests/unit/test_primitives.py tests/unit/test_api_progress.py` | 0 | 25 focused mutation, rollback, preservation, conflict, delegation, and adversarial tests passed. |
| `docker compose run --rm test pytest -p no:cacheprovider -q -m 'not integration' tests/unit` | 0 | 440 passed and 5 explicitly declared live/fixture skips. |
| `docker compose run --rm test mypy --cache-dir=/tmp/mypy-cache src/` | 0 | Strict typing passed for 89 source files. |
| `docker compose run --rm test ruff check --no-cache src tests` | 0 | Lint passed. |
| `docker compose run --rm test ruff format --no-cache --check .` | 0 | All 173 files were formatted. |
| `docker compose run --rm test bandit -r src -c pyproject.toml` | 0 | No security issues were identified. |
| `docker compose run --rm test odstool --help` | 0 | The installed CLI exposed all nine required commands. |

## Prompt 05 verification

Prompt 05 is implemented and verified locally and in remote Docker CI.
`migration-check`
loads strict versioned YAML/JSON metadata, environments, scenarios, actions,
observations, step results, traces, assertions, and evidence manifests.
Required actions and observations fail closed across all five execution states.
Typed values preserve empty cells, empty strings, zero, Booleans, and formula
errors; numeric comparisons use only declared absolute and relative tolerances.

The CLI exposes `run`, `inspect`, `diff`, `mutate`, and `report`. Acceptance
runs produce content-addressed JSON and Markdown evidence. Mutation campaigns
modify embedded Python and ODF formulas only in copied, transactionally
verified ODS packages; infrastructure unavailability is inconclusive rather
than a killed mutant. The checked-in YAML example treats save, close, and
reopen as explicit actions. Expectations are authored from requirements and
independently reviewed; cached Excel values are not an authoritative oracle.
Prompt 06 now routes target execution through the stateful runtime service.

| Exact command | Exit | Outcome |
|---|---:|---|
| `docker compose run --rm test pytest -q tests/unit/test_migration_check.py` | 0 | 9 focused schema, status, assertion, evidence, CLI, lifecycle, mutation, and report tests passed. |
| `docker compose run --rm test pytest -q tests/unit` | 0 | 450 passed and 5 explicitly declared live/fixture skips. |
| `docker compose run --rm test mypy src` | 0 | Strict typing passed for 93 source files. |
| `docker compose run --rm test ruff check src tests` | 0 | Lint passed. |
| `docker compose run --rm test ruff format --check .` | 0 | All 178 files were formatted. |
| `docker compose run --rm test bandit -q -r src -c pyproject.toml` | 0 | No security issues were identified. |
| `docker compose run --rm test python tools/ci_check.py package` | 0 | Wheel and sdist built; the packaged startup guard and metadata passed. |
| `docker compose run --rm test python -m xlsliberator.migration_check --help` | 0 | All five required commands were exposed from the current source. |

## Prompt 06 verification

Prompt 06 is implemented and verified in
[remote Docker CI run 29649373322](https://github.com/johannhartmann/xlsliberator/actions/runs/29649373322)
at commit `15a28ea`. The runtime boundary exposes exactly the 19 curated stateful
LibreOffice operations. Every operation requires an explicit session ID and
reports transport state separately from operation status. Sessions receive an
isolated LibreOffice profile, port, display, working copy, log directory, and
the exact pinned office build. Workspace-root policy, localhost-only transport,
timeouts, forced container cleanup, retained failure evidence, truthful
unavailable UI behavior, and disposable source-preserving document copies are
enforced by the service.

The migration checker and scenario runner use the stateful service rather than
starting LibreOffice directly. The stable entry point is
`xlsliberator libreoffice-mcp-serve`; the historical MCP command is a deprecated
wrapper. Unit coverage includes fake backend and fake client transport tests,
session correlation, isolation, cleanup, path rejection, timeout handling, and
scenario evidence. The blocking integration job exercises the real
LibreOffice 26.2.4.2 image entirely through Docker.

| Exact CI command | Exit | Outcome |
|---|---:|---|
| `docker run --rm --network none --read-only --tmpfs /tmp:rw,noexec,nosuid,size=1g,mode=1777 --volume "$PWD:/workspace" --workdir /workspace xlsliberator-test:py3.11 pytest -p no:cacheprovider -m "not integration" --junitxml=artifacts/pytest-unit-3.11.xml` | 0 | Python 3.11 office-free unit matrix passed. |
| `docker run --rm --network none --read-only --tmpfs /tmp:rw,noexec,nosuid,size=1g,mode=1777 --volume "$PWD:/workspace" --workdir /workspace xlsliberator-test:py3.12 pytest -p no:cacheprovider -m "not integration" --junitxml=artifacts/pytest-unit-3.12.xml` | 0 | Python 3.12 office-free unit matrix passed. |
| `docker compose run --rm test ruff format --no-cache --check .` | 0 | Formatting gate passed. |
| `docker compose run --rm test ruff check --no-cache .` | 0 | Lint gate passed. |
| `docker compose run --rm test mypy --cache-dir=/tmp/mypy-cache src/` | 0 | Strict typing gate passed. |
| `docker compose --profile ci-orchestrator run --rm test-orchestrator python tools/ci_check.py office` | 0 | Blocking disposable-office integration passed against LibreOffice 26.2.4.2 and uploaded retained evidence. |
| `docker compose --profile ci-orchestrator run --rm test-orchestrator python tools/ci_check.py docker-web` | 0 | Blocking Docker web smoke passed and uploaded its attestation and test report. |
| `docker compose --profile ci-orchestrator run --rm security-audit python tools/ci_check.py security` | 0 | Docker security gate passed. |
| `docker compose run --rm test python tools/ci_check.py package` | 0 | Wheel, sdist, packaged startup guard, and metadata checks passed. |

## Current architecture problems

| Problem | Current evidence | Required destination |
|---|---|---|
| Legacy provider code | deprecated modules remain available only under the optional `xlsliberator.legacy_agent` extra for migration compatibility | remove after all downstream users have moved to `xlsliberator-swe` |
| Core deterministic surface | typed primitives, `xlsprobe`, transactional `odstool`, fail-closed `migration-check`, and the stateful target runtime exist | Prompt 07 composes them from the thin Open-SWE fork |
| Fail-open or ambiguous validation states | macro/control gates can report `SKIPPED` for missing output; legacy translation fallbacks return source text; some test paths accept skipped live behavior | Prompt 01 makes required gates reject `SKIPPED`, `UNAVAILABLE`, `NOT_RUN`, timeouts, and transport-only success |
| Integration coverage is incomplete outside CI | GitHub office and web jobs are now blocking with `XLSLIBERATOR_FAIL_ON_SKIP=1`, but `make all` omits office, web, and package gates; multiple integration modules still contain permissive skips outside that mode | Prompt 01 aligns local and CI truth and separates explicitly optional live-provider tests |
| Placeholder or inaccurate tools | `take_screenshot` and keyboard input are registered MCP tools but only return unavailable; “click” resolves a handler and invokes a script rather than dispatching a GUI click; `test_placeholder.py` adds no behavior coverage | Prompt 01 corrects naming/status; Prompts 04 and 06 expose only operations with truthful semantics |
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
| 01 — truthful validation and CI | COMPLETE; REMOTE CI GREEN AT `9cce352` | fail-closed tests, aligned Docker CI commands, exact results |
| 02 — extract model orchestration | COMPLETE; REMOTE CI GREEN AT `9cce352` | dependency/import audit, removed prohibited runtime/worker paths, typed primitive tests |
| 03 — `xlsprobe` dossier | COMPLETE; REMOTE CI GREEN AT `9cce352` | CLI/API schema, fixture snapshots, dossier evidence |
| 04 — transactional `odstool` | COMPLETE; REMOTE CI GREEN AT `aac9a01` | mutation-plan, rollback, preservation and conflict tests |
| 05 — `migration-check` | COMPLETE; REMOTE CI GREEN AT `aac9a01` | scenario schema, deterministic target traces, fail-closed negatives, mutation evidence |
| 06 — stateful LibreOffice MCP | COMPLETE; REMOTE CI GREEN AT `15a28ea` | session lifecycle, isolation, runtime integration evidence |
| 07 — thin Open-SWE fork | COMPLETE IN `xlsliberator-swe`; REMOTE CI GREEN AT `27d2aa19` | separate repository, upstream record, sync procedure |
| 08 — sandbox snapshot | COMPLETE IN `xlsliberator-swe`; REMOTE CI GREEN AT `84d22454` | image/SBOM, tool versions, sandbox smoke |
| 09 — triggers and hydration | COMPLETE IN `xlsliberator-swe`; REMOTE CI GREEN AT `f513e485` | attachment tests, artifact hashes, durable-thread evidence |
| 10 — Deep Agents skills | COMPLETE ACROSS BOTH REPOSITORIES | skill discovery/loading tests, isolation proof and migration-only wiring |
| 11 — forensics/planning/testing/package skills | COMPLETE; REMOTE CI GREEN AT `0fc39e5` | skill fixtures and trajectory tests |
| 12 — migration-specialist skills | COMPLETE; REMOTE CI GREEN AT `0fc39e5` | formula/VBA/UI/dependency/LO skill tests |
| 13 — specialist subagents and routing | COMPLETE IN `xlsliberator-swe` AT `886a23dc` | routing tests, model config, isolated skill/tool scopes |
| 14 — curated MCP tools | COMPLETE IN `xlsliberator-swe`; REMOTE CI GREEN AT `37c52efa` | allowlist, typed failures, integration traces |
| 15 — deterministic middleware | COMPLETE IN `xlsliberator-swe` AT `49164719` | checkpoint and anti-fake-success/anti-test-weakening tests |
| 16 — migration lead | COMPLETE IN `xlsliberator-swe` AT `18e0d770` | workflow tests and end-to-end thread trajectory |
| 17 — independent reviewer | COMPLETE IN `xlsliberator-swe`; REMOTE CI GREEN AT `7c7ee4c8` | separate reviewer identity, hidden/mutation gate evidence; Agent and sandbox runs `29654989184` and `29654989178` |
| 18 — web to Open-SWE threads | IMPLEMENTED ACROSS BOTH REPOSITORIES; CI FINALIZATION IN PROGRESS | authenticated job/thread mapping, safe stages, resume, cancellation, owner-scoped artifact delivery and fake-service Docker smoke |
| 19 — serious demos and corpus | IMPLEMENTED; CI EVIDENCE PENDING | eight licensed episodes, behavioral public scenarios, corpus search/subsets, evidence-derived feature report |
| 20 — repair promotion/build farm | IMPLEMENTED IN `xlsliberator`; OPEN-SWE FINALIZATION PENDING | validated repair records, public/reviewer corpus MCP, fail-closed build-farm contract, real TDF-172479 differential |
| 21 — execution hardening | IMPLEMENTED IN `xlsliberator`; OPEN-SWE FINALIZATION PENDING | networkless/read-only runtime, bounded hostile-input parser, typed grants, escape probes, security-adversary evaluator, Bandit and dependency audit |
| 22 — LangSmith evaluations | PENDING | datasets, evaluators, thresholds, release gate |
| 23 — autonomous showcase | PENDING | reproducible full episode, target evidence, independent verdict |

## Next action

Obtain blocking remote CI evidence for Prompts 19–20, complete the Open-SWE
repair-promotion orchestration, then begin **Prompt 21 — hostile-workbook
hardening**. Local commands continue exclusively in Docker; no local Python or
office fallback is permitted.
