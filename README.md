# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

**An Open-SWE agent for Excel-to-LibreOffice migration**

XLSLiberator uses an embedded Open-SWE workflow to migrate Excel files
(`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` files. Open-SWE
inspects the workbook, plans and performs the migration, exercises the result in
the pinned LibreOffice Docker target, reviews the evidence, and produces the
deliverables. Formula, VBA, control, and runtime support varies by workbook and
must be read from the
[evidence-backed capability matrix](docs/capability_matrix.md).

## Features

- **Open-SWE Migration Agent**: Owns the migration thread, tool use, repairs,
  validation, review, and delivery
- **Formula Translation**: AST-based formula repair tools for Excel→Calc compatibility
- **Raw VBA Extraction**: Preserves complete source modules for the embedded Open-SWE workflow
- **Target-native Script Upsert**: Transactionally embeds caller-supplied Python/UNO modules
- **Translation Evidence**: Records syntax, export, provenance, runtime, and
  unresolved-artifact outcomes without promoting model confidence to success
- **Embedded Python Macros**: Embeds converted macros directly into the ODS file with event handling
- **Safe-by-Default Macros**: Never changes host macro security; runtime checks,
  when requested, run only in disposable Docker profiles
- **Validated Transformation**: Experimental certification pipeline for the pinned LibreOffice Docker runtime; current measured capabilities are listed in the [capability matrix](docs/capability_matrix.md)
- **Native LibreOffice Conversion**: Uses LibreOffice as the base converter; semantic equivalence requires scenario evidence
- **Artifact Inventory**: Detects supported and unsupported workbook structures without claiming universal coverage
- **MCP Server**: FastMCP server with explicit transport and operation status;
  unimplemented capabilities return `UNAVAILABLE`

## Prerequisites

### System Requirements

**Docker with the pinned LibreOffice 26.2.4.2 runtime image.** A host Python or
host office installation is neither required nor supported, including for
diagnostics.

Open-SWE is the only supported agent and orchestrator. This repository builds a
pinned upstream Open-SWE runtime plus its XLSLiberator graph and authenticated
workbook API as the internal `xlsliberator-open-swe` Compose service. There is
no second repository, maintained fork, alternate agent, competing migration
orchestrator, or local web fallback. The inspection, package-editing, and
LibreOffice interfaces are tools used by Open-SWE, not a second workflow.

A model and its matching credential must be selected explicitly. GitHub Models
is disabled by default and never acts as a fallback or gate. Model-free
inspection and validation tools do not read provider credentials.

## Installation

XLSLiberator is Docker-only. Do not install or run its Python package on the host.

```bash
git clone https://github.com/johannhartmann/xlsliberator.git
cd xlsliberator
docker compose build test libreoffice-runtime xlsliberator-open-swe xlsliberator-web
```

## Quick Start

### Open-SWE web workflow

```bash
mkdir -p artifacts/runtime-tmp
cp .env.example .env
# Select one supported model and set only its matching provider key in .env.
# Example: XLSLIBERATOR_OPEN_SWE_MODEL=openai:gpt-5.5
docker compose up -d --build xlsliberator-web
```

Open `http://127.0.0.1:8080/`, choose an example workbook or upload your own,
and start the migration. The web service creates an authenticated Open-SWE
thread and streams its progress. It does not run a local converter or fall back
to another orchestrator.

### Component command line

These Docker-only commands expose individual XLSLiberator tools for development,
diagnostics, and explicit component use. They are not an alternate migration
agent.

```bash
# Run the base LibreOffice conversion tool
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator convert "$PWD/input.xlsx" "$PWD/output.ods"

# Inventory VBA without asking Open-SWE to migrate it
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator convert --no-macros "$PWD/input.xlsm" "$PWD/output.ods"

# Convert and run validation gates, producing a certification report
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator transform-validated "$PWD/input.xlsm" "$PWD/output.ods"

# Inspect a workbook, or validate an existing conversion, without converting
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator inspect "$PWD/input.xlsm"
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator validate "$PWD/input.xlsm" "$PWD/output.ods"
```

### Workbook forensics

`xlsprobe` is the read-only, model-free source inspection CLI used to prepare
migrations. Run it only in the Docker application boundary:

```bash
# Create the complete migration dossier under artifacts/source-case/migration/
mkdir -p artifacts/source-case
docker compose --profile ci-runner run --rm test-runner \
  xlsprobe dossier "$PWD/input.xlsm" --output "$PWD/artifacts/source-case"

# Query individual evidence surfaces without creating a dossier
docker compose --profile ci-runner run --rm test-runner \
  xlsprobe package-tree "$PWD/input.xlsm"
docker compose --profile ci-runner run --rm test-runner \
  xlsprobe extract-vba "$PWD/input.xlsm"
docker compose --profile ci-runner run --rm test-runner \
  xlsprobe formulas "$PWD/input.xlsm"
```

The other focused commands are `inspect`, `controls`, `dependencies`, and
`previews`. Every command accepts `--timeout-seconds` and `--max-source-mib`.
Programmatic callers can apply the stricter `ProbeLimits` contract for archive
entry counts, per-entry and aggregate expansion size, compression ratio, and
preview size. Nested archives are retained as raw evidence and never expanded.

The versioned dossier contains a byte-for-byte `workbook.original`, bounded
metadata and sheet previews, formulas grouped by sheet or defined name, complete
extracted VBA module text and boundaries, controls, relationships, dependency
findings, raw package parts or OLE streams, and explicit coverage gaps. Its
`dossier.md` is an index with an untrusted-content boundary; large or
executable-looking workbook content is referenced instead of copied into the
model-readable markdown. Dossier creation snapshots the source transactionally,
refuses concurrent source changes, and never replaces an existing dossier.

### Transactional ODS package edits

`odstool` is the transactional package-editing boundary for scripts and event bindings.
It verifies the source package before editing, writes and verifies a complete
candidate beside the original, fsyncs it, rejects concurrent source changes,
and atomically replaces the source only after every check passes. Existing
scripts, unknown package members, ZIP metadata, and unrelated manifest entries
are preserved.

```bash
# Inspect and verify without changing the package
docker compose --profile ci-runner run --rm test-runner \
  odstool list "$PWD/output.ods"
docker compose --profile ci-runner run --rm test-runner \
  odstool verify "$PWD/output.ods"
docker compose --profile ci-runner run --rm test-runner \
  odstool inspect-scripts "$PWD/output.ods"

# Preview an upsert, then commit against the reviewed source hash
docker compose --profile ci-runner run --rm test-runner \
  odstool upsert-script "$PWD/output.ods" "$PWD/repair.py" --dry-run
docker compose --profile ci-runner run --rm test-runner \
  odstool upsert-script "$PWD/output.ods" "$PWD/repair.py" \
  --expect-sha256 <reviewed-package-sha256>

# Compare or snapshot verified packages
docker compose --profile ci-runner run --rm test-runner \
  odstool diff "$PWD/before.ods" "$PWD/after.ods"
docker compose --profile ci-runner run --rm test-runner \
  odstool snapshot "$PWD/output.ods" --output "$PWD/artifacts/output-snapshot"
```

The remaining mutation commands are `remove-script`, `bind-event`, and
`unbind-event`; each supports `--dry-run` and `--expect-sha256`. Event binding
YAML has the closed schema `id`, `control_id`, `event_name`, `module`, and
`function`. The target module and exported function must already resolve.
Mutation results include member-level diffs and explicitly report invalidated
package signatures. A failed validation, write, binding resolution, or
precondition leaves the original untouched.

### Public migration acceptance

`migration-check` loads a strict, versioned YAML or JSON contract containing
independently reviewed migration metadata, the declared environment and
capability grants, and an action/observation scenario. Required actions and
assertions fail closed: `UNAVAILABLE`, `SKIPPED`, and `NOT_RUN` can never
produce a passing acceptance result. Typed observations keep empty cells,
empty strings, numbers, Booleans, and formula errors distinct, with explicitly
declared absolute and relative numeric tolerances.

```bash
# Execute only through the trusted Docker test runner; LibreOffice remains in
# the pinned office worker image.
docker compose --profile ci-runner run --rm test-runner \
  migration-check run "$PWD/public-acceptance.yaml" "$PWD/output.ods" \
  --output "$PWD/artifacts/acceptance-evidence"

docker compose run --rm test \
  migration-check inspect "$PWD/artifacts/acceptance-evidence"
docker compose run --rm test \
  migration-check diff "$PWD/trace-a.json" "$PWD/trace-b.json"

# Mutants are generated only from copies under <migration-dir>/mutations.
docker compose --profile ci-runner run --rm test-runner \
  migration-check mutate "$PWD/migration"
docker compose run --rm test migration-check report "$PWD/migration"
```

Each run emits content-addressed JSON plus a Markdown report. Mutation campaigns
change one embedded Python module or ODF formula at a time and pass only when
the public acceptance detects every executable mutant. Runtime unavailability
is inconclusive, never a killed mutant. Expected results come from the authored
requirements and independent review; cached Excel values are explicitly not an
authoritative oracle. Real target execution uses the stateful session service,
which keeps the source immutable and executes every Office operation in the
pinned Docker worker image.

### Stateful LibreOffice MCP service

Start the trusted-local MCP server for Open-SWE or another explicit MCP client:

```bash
# Build the exact office worker, then start the loopback-only MCP gateway
mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
docker compose build libreoffice-runtime xlsliberator-mcp
docker compose up -d xlsliberator-mcp

# Client connects to: http://localhost:8080/mcp

# Equivalent one-shot server command inside the application image:
docker compose run --rm --service-ports xlsliberator-mcp \
  xlsliberator libreoffice-mcp-serve --host 0.0.0.0 --port 8000
```

The public surface contains exactly these session tools:

```text
create_session  open_document  inspect_document  list_sheets
read_cells  write_cells  list_formulas  recalculate  list_controls
dispatch_control_event  send_keyboard_event  execute_python_macro
capture_screenshot  export_pdf  save  close  reopen  collect_logs
destroy_session
```

`create_session` returns an explicit session ID; every other call requires it.
Each session has a unique profile identity, port/display, immutable runtime
identity, disposable working copy, and preserved logs/attachments. UI control,
keyboard, and screenshot operations return `UNAVAILABLE` unless a real event
layer can prove the operation; the service never invokes a handler and labels
that a click. The legacy `mcp-serve` command is deprecated, and legacy loose
tool names are not registered. See [`docs/mcp_tools.md`](docs/mcp_tools.md).

### Browser Web App

The [Quick Start](#open-swe-web-workflow) launches the complete browser-facing
stack. Pick a bundled example workbook or upload your own, start a real
migration, watch Open-SWE progress inline, and download the converted `.ods`
file plus JSON and Markdown reports.
Compose starts the internal MCP, Open-SWE, and web services in dependency order.
The web container is only an authenticated Open-SWE client. It fails closed
when the internal Open-SWE service is absent or unreachable and never invokes
conversion locally.

The supported development path is entirely Docker-based. The host shell is limited to Docker,
Git, and file operations; it must not start Python, `uv`, LibreOffice, UNO, or PyUNO.

The web app accepts `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` uploads. It stores each job
under a server-generated ID, uses isolated LibreOffice profiles for web conversions,
and avoids exposing internal filesystem paths in API responses.

See the [User Guide](user_guide.md) for the full workflow and the
[Web App Guide](docs/web_app.md) for development, API, and Docker details.

### Python API

The Python API is an in-container API. These examples run inside the application
Docker container; direct host invocation is rejected before any office runtime
can be constructed.

```python
from xlsliberator.api import convert

# Simple conversion
result = convert("input.xlsx", "output.ods")

# With options
result = convert(
    "input.xlsm",
    "output.ods",
    python_modules={"MigratedModule": generated_python_uno_source},
    embed_macros=True,
)

print(f"Conversion completed: {result.success}")
print(f"Total formulas: {result.total_formulas}")
print(f"VBA procedures: {result.vba_procedures}")
```

For the validated transformation pipeline (convert + certify against real office runtimes):

```python
from xlsliberator.validated_api import transform_validated

report = transform_validated("input.xlsm", "output.ods", targets=["libreoffice"])
print(f"Certified: {report.certification.certified}")
print(report.to_markdown())
```

## Python Macro Support

The low-level conversion API does not choose a model or translate VBA by itself.
The embedded Open-SWE workflow reads raw VBA and the workbook dossier, generates
target-native Python/UNO modules, inserts those artifacts transactionally, and
validates them in the LibreOffice target.

### How It Works

During conversion, XLSLiberator:
- Embeds Python-UNO macros into the ODS file's `Scripts/python/` directory
- Rewrites event handlers in `content.xml` from `language=Basic` to `language=Python`
- Loads documents only in the disposable pinned office worker container, where
  `MacroExecutionMode=4` is scoped to that isolated job profile
- Runs requested runtime macro checks inside isolated, temporary LibreOffice
  profiles; an unavailable or skipped check does not pass validation

**XLSLiberator does not modify global LibreOffice macro security.** The legacy
option is retained only for API compatibility and reports that global mutation
is unsupported:

```python
from xlsliberator.api import convert

convert("input.xlsm", "output.ods", allow_global_macro_security_change=True)
```

Do not configure or start a host LibreOffice installation. Macro execution uses
only an isolated, disposable profile inside the pinned office container.

## Architecture

Open-SWE owns the complete migration:

1. **Migration Thread**: the web app creates or resumes one authenticated,
   durable Open-SWE thread for the workbook
2. **Source Investigation**: Open-SWE uses the workbook dossier, formula, VBA,
   control, and dependency tools to understand the source
3. **Target Implementation**: Open-SWE plans the migration and produces the
   LibreOffice-compatible formulas, scripts, event bindings, and package changes
4. **LibreOffice Execution**: every office operation runs in the pinned
   LibreOffice 26.2.4.2 Docker target
5. **Evidence and Repair**: Open-SWE evaluates target behavior, performs bounded
   repairs, and records unresolved capabilities explicitly
6. **Independent Review and Delivery**: the migration is delivered only after
   the required review and acceptance evidence are present

The CLI, MCP methods, workbook inspection, package editing, and validation
modules are the tools behind this workflow. They do not constitute another
agent or orchestrator.

## Development

### Setup

```bash
# Clone repository
git clone https://github.com/yourusername/xlsliberator.git
cd xlsliberator

# Build the development and exact LibreOffice images
docker compose build test libreoffice-runtime

# Run tests
docker compose run --rm test pytest

# Code quality checks
make fmt      # Format code with ruff
make lint     # Lint with ruff
make typecheck # Type check with mypy
make test     # Run test suite
```

### Project Structure

```
xlsliberator/
├── src/xlsliberator/                # Main source code
│   ├── api.py                       # Base conversion and artifact application
│   ├── primitives.py                # Typed public operations
│   ├── xlsprobe.py                  # Bounded workbook forensics and dossiers
│   ├── validated_api.py             # Validated transformation (transform_validated)
│   ├── cli.py                       # Command-line interface
│   ├── config.py                    # Environment-driven configuration
│   ├── extract_excel.py             # Workbook metadata/IR extraction
│   ├── extract_vba.py               # VBA extraction (oletools)
│   ├── python_syntax_validator.py   # Phase 2: Syntax validation
│   ├── web/open_swe.py              # Authenticated Open-SWE transport
│   ├── embed_macros.py              # Macro embedding + event binding
│   ├── python_macro_manager.py      # Scripts/python/ management & validation
│   ├── formula_ast_transformer.py   # Formula AST transforms
│   ├── formula_engine.py            # Formula checks & rule registry
│   ├── fix_native_ods.py            # Post-conversion ODS fixes
│   ├── validation_runner.py         # Validation gate sequencing
│   ├── certification_report.py      # Certification report writer
│   ├── calc_backend.py              # Office backend discovery & isolated profiles
│   ├── control_inventory.py         # ODS form/control/event inventory
│   ├── uno_conn.py                  # In-process LibreOffice UNO connection
│   ├── lo_worker.py                 # Out-of-process UNO worker (stdlib-only)
│   ├── lo_worker_client.py          # Client for the UNO worker
│   ├── libreoffice_session.py       # Stateful Docker-only runtime sessions
│   ├── libreoffice_mcp.py           # Session-oriented MCP tools
│   ├── mcp_server.py                # Curated FastMCP service
│   └── web/                         # FastAPI web app
├── tests/                           # Test suite
│   ├── unit/                        # Unit tests
│   ├── it/                          # LibreOffice/UNO integration tests
│   ├── real/                        # Real-workbook fixtures
│   ├── bench/                       # Performance benchmarks
│   └── integration/                 # Docker web-app tests
├── rules/                           # Formula/event mapping rules (YAML)
└── docs/                            # Documentation
```

## Testing

```bash
# Run all tests in the development container
docker compose run --rm test pytest

# Run specific test categories
docker compose run --rm test pytest -m "not integration"
make test-integration
docker compose run --rm test pytest -m benchmark

# Run with coverage
docker compose run --rm test pytest --cov=xlsliberator --cov-report=html
```

## Measured capabilities

No repository-wide equivalence or VBA-success percentage is currently certified.
Measured results are generated from the conformance corpus and published in the
[capability matrix](docs/capability_matrix.md) and generated
[release-readiness report](docs/release_readiness.md); unavailable, skipped,
unsupported, waived, and failed runs remain distinct from passing results.
Tool-level capability results do not by themselves prove a complete Open-SWE
migration. Agent acceptance evidence is tracked separately. See the
[agentic implementation ledger](docs/agentic_implementation_status.md) and the
[Open-SWE architecture](docs/architecture/open-swe-migration.md).

## Known Limitations

- Core conversion does not translate VBA; Open-SWE must supply target-native modules
- Complex migrations require Open-SWE specialist work and independent review
- Cross-workbook references require manual adjustment
- COM automation and external DLLs are not supported

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Ensure all tests pass (`make test`)
5. Run code quality checks (`make fmt lint typecheck`)
6. Submit a pull request

## License

This project is licensed under the **GNU General Public License v3.0 or later (GPLv3+)**.

See [LICENSE](LICENSE) for the full license text.

## Support

- **Issues**: [GitHub Issues](https://github.com/johannhartmann/xlsliberator/issues)
- **Discussions**: [GitHub Discussions](https://github.com/johannhartmann/xlsliberator/discussions)

## Author

**Johann-Peter Hartmann**
Email: johann-peter.hartmann@mayflower.de
GitHub: [@johannhartmann](https://github.com/johannhartmann)

## Acknowledgments

- **Lukas Kahwe Smith** ([@lsmith77](https://github.com/lsmith77)): For the original idea and concept
- **LibreOffice**: For the excellent open-source office suite
- **oletools**: For VBA extraction capabilities

## Roadmap

- [x] Embed pinned upstream Open-SWE in the single-repository Docker stack
- [ ] Record explicitly authorized live-model acceptance evidence
- [ ] Enhanced formula repair logic
- [ ] Integration tests for VBA quality pipeline
