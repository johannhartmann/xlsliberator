# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

**Deterministic Excel-to-LibreOffice migration toolbelt**

XLSLiberator experimentally converts Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` files. Formula, VBA, control, and runtime support varies by artifact and must be read from the [evidence-backed capability matrix](docs/capability_matrix.md).

## Features

- **Formula Translation**: Deterministic AST-based formula transformation for Excel→Calc compatibility
- **Raw VBA Extraction**: Preserves complete source modules for external migration agents
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

No model-provider credential is read by deterministic commands. Long-running
model orchestration, model selection, and credentials belong to the separate
`xlsliberator-swe` Open-SWE service. The deprecated embedded translator can be
installed only with the explicit `legacy-agent` extra and is disabled by default.

## Installation

XLSLiberator is Docker-only. Do not install or run its Python package on the host.

```bash
git clone https://github.com/johannhartmann/xlsliberator.git
cd xlsliberator
docker compose build test libreoffice-runtime
```

## Quick Start

### Command Line

```bash
# Basic conversion from the Docker application orchestrator
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator convert "$PWD/input.xlsx" "$PWD/output.ods"

# VBA is inventoried but not silently translated by the deterministic converter
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator convert --no-macros "$PWD/input.xlsm" "$PWD/output.ods"

# Convert and run validation gates, producing a certification report
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator transform-validated "$PWD/input.xlsm" "$PWD/output.ods"

# Inspect a workbook, or validate an existing conversion, without converting
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator inspect "$PWD/input.xlsm"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator validate "$PWD/input.xlsm" "$PWD/output.ods"
```

### Workbook forensics

`xlsprobe` is the read-only, model-free source inspection CLI used to prepare
migrations. Run it only in the Docker application boundary:

```bash
# Create the complete migration dossier under artifacts/source-case/migration/
mkdir -p artifacts/source-case
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsprobe dossier "$PWD/input.xlsm" --output "$PWD/artifacts/source-case"

# Query individual evidence surfaces without creating a dossier
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsprobe package-tree "$PWD/input.xlsm"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsprobe extract-vba "$PWD/input.xlsm"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
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

`odstool` is the deterministic mutation boundary for scripts and event bindings.
It verifies the source package before editing, writes and verifies a complete
candidate beside the original, fsyncs it, rejects concurrent source changes,
and atomically replaces the source only after every check passes. Existing
scripts, unknown package members, ZIP metadata, and unrelated manifest entries
are preserved.

```bash
# Inspect and verify without changing the package
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool list "$PWD/output.ods"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool verify "$PWD/output.ods"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool inspect-scripts "$PWD/output.ods"

# Preview an upsert, then commit against the reviewed source hash
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool upsert-script "$PWD/output.ods" "$PWD/repair.py" --dry-run
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool upsert-script "$PWD/output.ods" "$PWD/repair.py" \
  --expect-sha256 <reviewed-package-sha256>

# Compare or snapshot verified packages
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool diff "$PWD/before.ods" "$PWD/after.ods"
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  odstool snapshot "$PWD/output.ods" --output "$PWD/artifacts/output-snapshot"
```

The remaining mutation commands are `remove-script`, `bind-event`, and
`unbind-event`; each supports `--dry-run` and `--expect-sha256`. Event binding
YAML has the closed schema `id`, `control_id`, `event_name`, `module`, and
`function`. The target module and exported function must already resolve.
Mutation results include member-level diffs and explicitly report invalidated
package signatures. A failed validation, write, binding resolution, or
precondition leaves the original untouched.

### Provider-neutral MCP Server

Start the MCP server for any MCP-compatible orchestrator:

```bash
# Build the exact office worker, then start the loopback-only MCP orchestrator
mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
docker compose build libreoffice-runtime xlsliberator-mcp
docker compose up -d xlsliberator-mcp

# Client connects to: http://localhost:8080/mcp
```

**Available Tools:**
- `convert_excel_to_ods` - Run deterministic native Excel-to-ODS conversion
- `inspect_workbook` - Return parsed workbook inventory and unsupported artifacts
- `validate_transformation` - Run validation gates and return certification data
- `compare_formulas` - Test formula equivalence
- `read_cell`, `list_sheets`, `get_sheet_data` - Read spreadsheet data
- `list_controls`, `list_event_bindings` - Inspect ODS form controls and event bindings
- `embed_macros`, `validate_macros`, `list_embedded_macros`, `test_macro_execution` - Manage and test Python-UNO macros
- `recalculate_document`, `validate_document_runtime` - Run target operations in a disposable pinned Docker runtime
- `execute_button_handler` - Resolve an inventoried button and invoke its handler directly; this is not a GUI click
- `open_document_gui`, `click_form_button`, `send_keyboard_input`, `take_screenshot` - Explicitly unavailable until real container capabilities exist
- `get_cell_colors` - Inspect cell background state through the Docker worker

All tools are registered in `src/xlsliberator/mcp_server.py`. Their canonical
response and runtime-selection contracts are documented in
[`docs/mcp_tools.md`](docs/mcp_tools.md).

### Browser Web App

Run the Docker web app for the closest production-like setup:

```bash
mkdir -p artifacts/runtime-tmp
docker compose build libreoffice-runtime xlsliberator-web
docker compose up -d xlsliberator-web
```

Open `http://127.0.0.1:8080/` for the landing page and its embedded live demo: pick a
bundled example workbook (or upload your own), start a real conversion, watch the pipeline
progress inline, and download the converted `.ods` file plus JSON and Markdown reports.

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

XLSLiberator does not choose a model or translate VBA in its deterministic
conversion API. External agents read the raw VBA and workbook dossier, generate
target-native Python/UNO modules, and pass those modules back for transactional
upsert and independent validation.

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

XLSLiberator uses a deterministic target-tool approach:

1. **Native Conversion**: the pinned LibreOffice Docker runtime provides the base conversion; equivalence is evaluated separately
2. **Source Inspection**: extracts workbook metadata and raw VBA without model calls
3. **External Migration**: `xlsliberator-swe` owns model routing and specialist work
4. **Artifact Upsert**: embeds explicitly supplied target-native Python/UNO modules
5. **Isolated Runtime Validation**: runs required validation in disposable, resource-limited LibreOffice containers and profiles
6. **Formula Repair**: deterministic transformations address known incompatibilities

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
│   ├── api.py                       # Deterministic conversion pipeline
│   ├── primitives.py                # Typed public operations
│   ├── xlsprobe.py                  # Bounded workbook forensics and dossiers
│   ├── validated_api.py             # Validated transformation (transform_validated)
│   ├── cli.py                       # Command-line interface
│   ├── config.py                    # Environment-driven configuration
│   ├── extract_excel.py             # Workbook metadata/IR extraction
│   ├── extract_vba.py               # VBA extraction (oletools)
│   ├── python_syntax_validator.py   # Phase 2: Syntax validation
│   ├── legacy_agent/                # Deprecated optional provider-backed code
│   ├── embed_macros.py              # Macro embedding + event binding
│   ├── python_macro_manager.py      # Scripts/python/ management & validation
│   ├── formula_ast_transformer.py   # Formula AST transforms
│   ├── formula_engine.py            # Formula checks & rule registry
│   ├── fix_native_ods.py            # Post-conversion ODS fixes
│   ├── validation_runner.py         # Validation gate orchestration
│   ├── certification_report.py      # Certification report writer
│   ├── calc_backend.py              # Office backend discovery & isolated profiles
│   ├── control_inventory.py         # ODS form/control/event inventory
│   ├── uno_conn.py                  # In-process LibreOffice UNO connection
│   ├── lo_worker.py                 # Out-of-process UNO worker (stdlib-only)
│   ├── lo_worker_client.py          # Client for the UNO worker
│   ├── mcp_server.py / mcp_tools.py # FastMCP server and tools
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
The checked-in generated readiness report belongs to the earlier certification
architecture. It is not evidence that the Open-SWE autonomous migration system
is complete. The current Docker-backed baseline is blocked by a Docker Desktop
storage I/O failure; see the
[agentic implementation ledger](docs/agentic_implementation_status.md).

## Known Limitations

- Core conversion does not translate VBA; orchestration must supply target-native modules
- Complex migrations require external specialist work and independent review
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

- [ ] Migrate model orchestration and independent review to `xlsliberator-swe`
- [ ] Enhanced formula repair logic
- [ ] Integration tests for VBA quality pipeline
