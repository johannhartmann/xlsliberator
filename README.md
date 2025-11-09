# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

**Excel to LibreOffice Calc converter with VBA-to-Python-UNO macro translation**

XLSLiberator converts Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` files with full formula translation and VBA-to-Python-UNO macro conversion.

## Features

- **Formula Translation**: Deterministic AST-based formula transformation for Excel→Calc compatibility
- **VBA-to-Python-UNO Conversion**: Translates Excel VBA macros to Python-UNO scripts
- **Embedded Python Macros**: Embeds converted macros directly into the ODS file with event handling
- **Native LibreOffice Conversion**: Uses LibreOffice's native conversion engine for 100% formula equivalence
- **Comprehensive Support**: Handles tables, charts, forms, and complex workbook structures
- **High Performance**: Processes 27k+ cells in under 5 minutes

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
# Basic conversion
xlsliberator convert input.xlsx output.ods

# With VBA macro translation
export ANTHROPIC_API_KEY="your-api-key"
xlsliberator convert --translate-vba input.xlsm output.ods

# Batch conversion
xlsliberator batch input_folder/ output_folder/
```

### Python API

```python
from xlsliberator import convert_excel_to_ods

# Simple conversion
result = convert_excel_to_ods("input.xlsx", "output.ods")

# With options
result = convert_excel_to_ods(
    "input.xlsm",
    "output.ods",
    translate_vba=True,
    embed_macros=True,
    repair_formulas=True
)

print(f"Conversion completed: {result.success}")
print(f"Formulas translated: {result.formula_count}")
print(f"VBA macros converted: {result.macro_count}")
```

## Architecture

XLSLiberator uses a hybrid approach:

1. **Native Conversion**: LibreOffice's native `--convert-to ods` provides the base conversion with 100% formula equivalence
2. **VBA Extraction**: Extracts VBA code from Excel files using oletools
3. **LLM Translation**: Translates VBA to Python-UNO using Claude API
4. **Macro Embedding**: Embeds translated Python macros into the ODS file via UNO
5. **Formula Repair**: Deterministic AST transformations fix incompatibilities

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
├── src/xlsliberator/         # Main source code
│   ├── api.py                # High-level API
│   ├── cli.py                # Command-line interface
│   ├── extract_vba.py        # VBA extraction
│   ├── vba2py_uno.py         # VBA→Python translation
│   ├── embed_macros.py       # Macro embedding
│   ├── formula_ast_transformer.py  # Formula repair
│   ├── fix_native_ods.py     # Post-conversion fixes
│   └── uno_conn.py           # LibreOffice UNO connection
├── tests/                    # Test suite
│   ├── unit/                 # Unit tests
│   ├── it/                   # Integration tests
│   └── data/                 # Test fixtures
├── rules/                    # Formula transformation rules
└── docs/                     # Documentation
```

## Testing

```bash
# Run all tests
pytest

# Run specific test categories
pytest -m unit           # Unit tests only
pytest -m integration    # Integration tests (requires LibreOffice)
pytest -m benchmark      # Performance benchmarks

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

- [ ] Support for more VBA patterns and validation
- [ ] Enhanced formula repair logic
