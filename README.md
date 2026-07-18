# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

**Excel to LibreOffice Calc converter with VBA-to-Python-UNO macro translation**

XLSLiberator experimentally converts Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` files. Formula, VBA, control, and runtime support varies by artifact and must be read from the [evidence-backed capability matrix](docs/capability_matrix.md).

## Features

- **Formula Translation**: Deterministic AST-based formula transformation for Excel→Calc compatibility
- **Experimental VBA-to-Python-UNO Conversion**: The legacy provider-backed path
  is not accepted unless deterministic validation and required target evidence pass
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

### Optional Requirements

**Anthropic API Key** (required only for VBA-to-Python translation)

1. Sign up at [anthropic.com](https://www.anthropic.com/)
2. Generate API key from console
3. Put `ANTHROPIC_API_KEY=your-api-key-here` in the untracked Compose `.env`
   file or supply it through the container platform's secret mechanism.

Without the API key, XLSLiberator can still attempt Excel-to-ODS conversion, but
VBA macros are not translated and formula preservation is not certified without
the required runtime and differential evidence.

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

# VBA translation requires ANTHROPIC_API_KEY in the Docker Compose environment
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  xlsliberator convert "$PWD/input.xlsm" "$PWD/output.ods"

# Skip VBA macro translation
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

### MCP Server (Claude Agent SDK)

Start the MCP server for Claude Agent SDK integration:

```bash
# Build the exact office worker, then start the loopback-only MCP orchestrator
mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
docker compose build libreoffice-runtime xlsliberator-mcp
docker compose up -d xlsliberator-mcp

# Client connects to: http://localhost:8080/mcp
```

Use with Claude Agent SDK:

```typescript
import { query } from "@anthropic-ai/claude-agent-sdk";

const mcpServers = {
  "libreoffice-uno": {
    url: "http://localhost:8080/mcp"
  }
};

for await (const message of query({
  prompt: generateMessages(),
  options: {
    mcpServers,
    allowedTools: ["mcp__libreoffice-uno__convert_excel_to_ods"],
  }
})) {
  // Agent can now convert Excel files!
}
```

**Available Tools:**
- `convert_excel_to_ods` - Convert Excel to ODS with macro translation
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
    embed_macros=True,   # translate and embed VBA macros (default)
    use_agent=True,      # legacy compatibility option; not certification evidence
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

When converting Excel files with VBA, XLSLiberator:

1. **VBA Translation**: Translates VBA macros to Python-UNO equivalents
2. **Event Handler Rewriting**: Rewrites VBA event handlers to point to the generated Python functions

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

### Configuring macro security in your LibreOffice

**Option 1: GUI Configuration**
- Open LibreOffice Calc
- Navigate to: `Tools → Options → LibreOffice → Security → Macro Security`
- Select **"Low"** (run all macros) or **"Medium"** (prompt for approval)

**Option 2: Trusted File Locations**
- Navigate to: `Tools → Options → LibreOffice → Security → Macro Security → Trusted Sources → Trusted File Locations`
- Add the directory containing your converted ODS files

## Architecture

XLSLiberator uses a hybrid approach:

1. **Native Conversion**: the pinned LibreOffice Docker runtime provides the base conversion; equivalence is evaluated separately
2. **VBA Extraction**: Extracts VBA code from Excel files using oletools
3. **Legacy LLM Translation**: Proposes VBA-to-Python-UNO candidates using the
   currently configured provider:
   - Phase 1: Reference-aware translation (hybrid LLM + regex pattern detection)
   - Phase 2: Python-UNO syntax validation (AST parsing, compilation checks)
   - Phase 3: legacy reflection and iterative refinement
   - Phase 4: target runtime evidence when available
4. **Macro Embedding**: Embeds translated Python macros into the ODS file via UNO
5. **Event Handler Rewriting**: Updates VBA event handlers to point to Python functions
6. **Isolated Runtime Validation**: Runs required validation in disposable, resource-limited LibreOffice containers and profiles
7. **Formula Repair**: Deterministic AST transformations fix incompatibilities

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
│   ├── api.py                       # Hybrid conversion pipeline (convert)
│   ├── validated_api.py             # Validated transformation (transform_validated)
│   ├── cli.py                       # Command-line interface
│   ├── config.py                    # Environment-driven configuration
│   ├── extract_excel.py             # Workbook metadata/IR extraction
│   ├── extract_vba.py               # VBA extraction (oletools)
│   ├── vba2py_uno.py                # VBA→Python-UNO translation entry point
│   ├── llm_vba_translator.py        # LLM-based VBA translator
│   ├── vba_reference_analyzer.py    # Phase 1: Reference-aware analysis
│   ├── python_syntax_validator.py   # Phase 2: Syntax validation
│   ├── vba_translation_validator.py # Phase 3: Quality evaluation
│   ├── vba_test_generator.py        # Phase 4: Test generation
│   ├── agent_rewriter.py            # Multi-agent rewriting for complex VBA
│   ├── embed_macros.py              # Macro embedding + event binding
│   ├── python_macro_manager.py      # Scripts/python/ management & validation
│   ├── runtime/                     # Excel-compatibility runtime for translated macros
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

- VBA translation requires Anthropic API key (Claude model)
- Some complex VBA patterns may require manual review
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
- **Anthropic**: For the Claude API used in VBA translation
- **oletools**: For VBA extraction capabilities

## Roadmap

- [ ] Migrate model orchestration and independent review to `xlsliberator-swe`
- [ ] Enhanced formula repair logic
- [ ] Integration tests for VBA quality pipeline
