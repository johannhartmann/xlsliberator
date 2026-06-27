# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

**Excel to LibreOffice Calc converter with VBA-to-Python-UNO macro translation**

XLSLiberator converts Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` files with full formula translation and VBA-to-Python-UNO macro conversion.

## Features

- **Formula Translation**: Deterministic AST-based formula transformation for Excel→Calc compatibility
- **VBA-to-Python-UNO Conversion**: Translates Excel VBA macros to Python-UNO scripts with 4-phase quality pipeline
- **Translation Quality Assurance**: Reference-aware translation, syntax validation, reflection loop, and automated test generation
- **Embedded Python Macros**: Embeds converted macros directly into the ODS file with event handling
- **Safe-by-Default Macros**: Embeds and validates Python macros in isolated LibreOffice profiles without changing your global macro security settings
- **Validated Transformation**: Optional certification pipeline that evaluates the output in LibreOffice/Apache OpenOffice and emits JSON + Markdown reports
- **Native LibreOffice Conversion**: Uses LibreOffice's native conversion engine for 100% formula equivalence
- **Comprehensive Support**: Handles tables, charts, forms, and complex workbook structures
- **High Performance**: Processes 27k+ cells in under 5 minutes
- **🆕 MCP Server**: FastMCP 2.0 server for Claude Agent SDK integration with 19 tools

## Prerequisites

### System Requirements

**LibreOffice 7.x+ with Python UNO bridge**

Ubuntu/Debian:
```bash
sudo apt-get update
sudo apt-get install libreoffice libreoffice-script-provider-python
```

Fedora/RHEL:
```bash
sudo dnf install libreoffice libreoffice-pyuno
```

macOS (Homebrew):
```bash
brew install --cask libreoffice
```

Windows:
- Download from [libreoffice.org](https://www.libreoffice.org/download/download/)
- Ensure Python support is included during installation

**Verify LibreOffice installation:**
```bash
soffice --version
```

**Python 3.11+**

Ubuntu/Debian:
```bash
sudo apt-get install python3.11 python3.11-venv python3-pip
```

macOS/Windows:
- Download from [python.org](https://www.python.org/downloads/)

### Optional Requirements

**Anthropic API Key** (required only for VBA-to-Python translation)

1. Sign up at [anthropic.com](https://www.anthropic.com/)
2. Generate API key from console
3. Set environment variable:
```bash
export ANTHROPIC_API_KEY="your-api-key-here"
```

Without the API key, XLSLiberator can still convert Excel to ODS with full formula preservation, but VBA macros will not be translated.

## Installation

### From Git

```bash
pip install git+https://github.com/johannhartmann/xlsliberator.git
```

### Development Installation

```bash
git clone https://github.com/johannhartmann/xlsliberator.git
cd xlsliberator
pip install -e ".[dev]"
```

## Quick Start

### Command Line

```bash
# Basic conversion (VBA macros are translated and embedded automatically when present)
xlsliberator convert input.xlsx output.ods

# VBA translation requires an Anthropic API key
export ANTHROPIC_API_KEY="your-api-key"
xlsliberator convert input.xlsm output.ods

# Skip VBA macro translation
xlsliberator convert --no-macros input.xlsm output.ods

# Convert and run validation gates, producing a certification report
xlsliberator transform-validated input.xlsm output.ods

# Inspect a workbook, or validate an existing conversion, without converting
xlsliberator inspect input.xlsm
xlsliberator validate input.xlsm output.ods
```

### MCP Server (Claude Agent SDK)

Start the MCP server for Claude Agent SDK integration:

```bash
# Start server with defaults (0.0.0.0:8000)
xlsliberator mcp-serve

# Client connects to: http://localhost:8000/mcp
```

Use with Claude Agent SDK:

```typescript
import { query } from "@anthropic-ai/claude-agent-sdk";

const mcpServers = {
  "libreoffice-uno": {
    url: "http://localhost:8000/mcp"
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
- `recalculate_document` - Force recalculation of all formulas
- `open_document_gui`, `click_form_button`, `send_keyboard_input`, `get_cell_colors`, `take_screenshot` - Drive GUI validation

All tools are registered in `src/xlsliberator/mcp_server.py`.

### Browser Web App

Run the Docker web app for the closest production-like setup:

```bash
docker compose up --build
```

Open `http://127.0.0.1:8080/`, upload an Excel workbook, watch the job progress page,
and download the converted `.ods` file plus JSON and Markdown reports.

For local development without Docker, install the optional web dependencies and run the
FastAPI app:

```bash
pip install -e ".[web,dev]"
xlsliberator web-serve --host 0.0.0.0 --port 8080 --reload
```

The web app accepts `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` uploads. It stores each job
under a server-generated ID, uses isolated LibreOffice profiles for web conversions,
and avoids exposing internal filesystem paths in API responses.

See the [User Guide](user_guide.md) for the full workflow and the
[Web App Guide](docs/web_app.md) for development, API, and Docker details.

### Python API

```python
from xlsliberator.api import convert

# Simple conversion
result = convert("input.xlsx", "output.ods")

# With options
result = convert(
    "input.xlsm",
    "output.ods",
    embed_macros=True,   # translate and embed VBA macros (default)
    use_agent=True,      # multi-agent rewriting for complex VBA (default)
)

print(f"Conversion completed: {result.success}")
print(f"Total formulas: {result.total_formulas}")
print(f"VBA procedures: {result.vba_procedures}")
```

For the validated transformation pipeline (convert + certify against real office runtimes):

```python
from xlsliberator.validated_api import transform_validated

report = transform_validated("input.xlsm", "output.ods", targets=["both"])
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
- Loads documents in-process with `MacroExecutionMode=4` (ALWAYS_EXECUTE_NO_WARN) so embedded macros run during conversion and validation
- Runs runtime macro validation inside isolated, temporary LibreOffice profiles

**Your global LibreOffice macro security is not modified by default.** To actually run the embedded
macros after opening the converted file in your own LibreOffice, configure macro security yourself (see
below). The legacy behavior that lowers the *global* macro security level is opt-in only:

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

1. **Native Conversion**: LibreOffice's native `--convert-to ods` provides the base conversion with 100% formula equivalence
2. **VBA Extraction**: Extracts VBA code from Excel files using oletools
3. **LLM Translation**: Translates VBA to Python-UNO using Claude API with 4-phase quality pipeline:
   - Phase 1: Reference-aware translation (hybrid LLM + regex pattern detection)
   - Phase 2: Python-UNO syntax validation (AST parsing, compilation checks)
   - Phase 3: Agentic reflection loop (self-evaluation and iterative refinement)
   - Phase 4: Runtime execution testing (UNO script execution validation)
4. **Macro Embedding**: Embeds translated Python macros into the ODS file via UNO
5. **Event Handler Rewriting**: Updates VBA event handlers to point to Python functions
6. **Isolated Runtime Validation**: Validates embedded macros in temporary LibreOffice profiles, leaving global macro security untouched
7. **Formula Repair**: Deterministic AST transformations fix incompatibilities

## Development

### Setup

```bash
# Clone repository
git clone https://github.com/yourusername/xlsliberator.git
cd xlsliberator

# Install dependencies
pip install -e ".[dev]"

# Run tests
pytest

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
# Run all tests
pytest

# Run specific test categories
pytest -m "not integration"   # Unit tests only (no LibreOffice required)
pytest -m integration         # Integration tests (requires LibreOffice)
pytest -m benchmark           # Performance benchmarks
LO_SKIP_IT=1 pytest           # Force-disable LibreOffice integration tests

# Run with coverage
pytest --cov=xlsliberator --cov-report=html
```

## Performance

Benchmark on real-world Excel file (27k cells, complex formulas):

- **Conversion time**: 264 seconds (< 5 minutes)
- **Formula equivalence**: 100% (using native conversion)
- **VBA translation**: 90%+ success rate
- **Memory usage**: < 500MB peak

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

- [x] Support for more VBA patterns and validation (4-phase quality pipeline)
- [ ] Enhanced formula repair logic
- [ ] Integration tests for VBA quality pipeline
