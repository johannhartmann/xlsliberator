# XLSLiberator

Excel to LibreOffice Calc converter with VBA-to-Python-UNO macro translation.

## Overview

**xlsliberator** converts Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`) to LibreOffice Calc `.ods` format with:

- ✅ Full formula translation with locale support (de-DE, en-US)
- ✅ VBA-to-Python-UNO macro conversion
- ✅ Embedded Python macros with event handling
- ✅ Tables, Charts, and Forms support
- ✅ Named ranges and structured references

## Installation

### Requirements

- Python 3.11+
- LibreOffice 7.x or 24.x (for conversion operations)
- conda environment `xlsliberator` (recommended)

### Setup

```bash
# Clone the repository
git clone <repository-url>
cd xlsliberator

# Install with uv (recommended)
uv pip install -e ".[dev]"

# Or with pip
pip install -e ".[dev]"
```

## Usage

### Command Line

```bash
# Basic conversion
xlsliberator convert input.xlsm output.ods

# With locale specification
xlsliberator convert input.xlsm output.ods --locale de-DE

# Strict mode (fail on unsupported features)
xlsliberator convert input.xlsm output.ods --strict

# With fallback import
xlsliberator convert input.xlsm output.ods --allow-fallback
```

### Python API

```python
from xlsliberator.api import convert

# Convert with options
report = convert(
    input_path="input.xlsm",
    output_path="output.ods",
    locale="de-DE",
    strict=False,
    enable_charts=True,
    enable_forms=True
)

# Check conversion report
print(f"Formulas translated: {report.formulas_translated}/{report.formulas_total}")
print(f"Unsupported features: {report.unsupported}")
```

## Development

### Project Structure

```
xlsliberator/
├── src/xlsliberator/      # Main source code
├── tests/                 # Test suite
│   ├── unit/             # Unit tests
│   ├── it/               # Integration tests
│   ├── bench/            # Performance benchmarks
│   └── real/             # Real dataset tests
├── rules/                # Formula and VBA mapping rules
├── docs/                 # Documentation
└── prompts/              # Implementation phase prompts
```

### Quality Checks

```bash
# Run all quality checks
make check

# Individual checks
make fmt        # Format code
make lint       # Lint code
make typecheck  # Type check
make test       # Run tests
```

### Running LibreOffice Integration Tests

```bash
# Start LibreOffice in headless mode
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &

# Run integration tests
pytest tests/it/ -q

# Skip integration tests
LO_SKIP_IT=1 pytest
```

## Implementation Status

See `prompts/checklist.md` for detailed implementation progress.

### Completed Phases
- [x] Phase 0.2: Feasibility Plan & Roadmap
- [ ] Phase 0.1: Repository Skeleton (in progress)

### Planned Features
- Formula mapper with 50+ functions
- VBA translator for common patterns
- Table and chart conversion
- Performance optimization
- Real dataset validation

## Testing

```bash
# Run all tests
pytest

# Run specific test categories
pytest tests/unit/          # Unit tests
pytest tests/it/            # Integration tests (requires LibreOffice)
pytest tests/bench/         # Benchmarks
pytest -m "not slow"        # Skip slow tests

# With coverage
pytest --cov=xlsliberator --cov-report=html
```

## Documentation

- [Feasibility Plan](docs/feasibility_plan.md) - Roadmap and milestones
- [Quality Gates](docs/gates.md) - Measurable success criteria
- [Implementation Phases](prompts/phases/) - Step-by-step prompts
- [Project Context](CLAUDE.md) - Development guidelines

## Architecture

### Data Flow

1. **Excel Ingestion** - Parse Excel files into intermediate representation (IR)
2. **VBA Extraction** - Extract and analyze VBA code modules
3. **Formula Mapping** - Translate Excel formulas to Calc equivalents
4. **ODS Generation** - Create LibreOffice Calc documents via UNO
5. **Macro Translation** - Convert VBA to Python-UNO code
6. **Embedding** - Inject Python macros into ODS files

### Key Components

- **IR Models** (Pydantic) - Neutral data representation
- **Formula Engine** - Tokenizer and rule-based translator
- **UNO Bridge** - LibreOffice headless connection
- **VBA Translator** - AST-based code generator
- **Report Generator** - Conversion metrics and warnings

## Performance Targets

| Operation | Target |
|-----------|--------|
| Excel Ingestion | ≥50k cells/min |
| Formula Mapping | ≥10k formulas/min |
| Full Conversion | <5 min/file |
| Memory Peak | <2 GB/file |

## Security

- VBA is analyzed **statically only** - no runtime execution
- No credential harvesting or malicious code generation
- Excel COM validator runs only in isolated sandbox

## Contributing

This project follows a phased implementation approach. See `prompts/phases/` for detailed implementation guides.

### Workflow

1. Follow phase order (F0 → F17)
2. Complete quality gates before proceeding
3. Update `prompts/checklist.md` after each phase
4. Commit after completing each phase

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Version

Current version: **0.1.0** (Alpha)

---

**Status:** 🚧 In Active Development

For questions and issues, please refer to the project documentation or contact the development team.
