# Contributing to XLSLiberator

Thank you for your interest in contributing to XLSLiberator! This document provides guidelines and instructions for contributing.

## Code of Conduct

Please be respectful and constructive in all interactions. We're here to build great software together.

## Getting Started

### Development Setup

1. **Fork and clone the repository**
   ```bash
   git clone https://github.com/johannhartmann/xlsliberator.git
   cd xlsliberator
   ```

2. **Install dependencies**
   ```bash
   pip install -e ".[dev]"
   ```

3. **Install LibreOffice** (for integration tests)
   ```bash
   # Ubuntu/Debian
   sudo apt-get install libreoffice libreoffice-calc python3-uno

   # macOS
   brew install --cask libreoffice
   ```

4. **Set up pre-commit hooks** (optional but recommended)
   ```bash
   pip install pre-commit
   pre-commit install
   ```

## Development Workflow

###  1. Create a branch
```bash
git checkout -b feature/your-feature-name
```

### 2. Make your changes
- Write code following the style guide below
- Add tests for new functionality
- Update documentation as needed

### 3. Run quality checks
```bash
# Format code
ruff format .

# Lint
ruff check .

# Type check
mypy src/

# Run tests
pytest
```

### 4. Commit your changes
```bash
git add .
git commit -m "feat: add feature description"
```

We use [Conventional Commits](https://www.conventionalcommits.org/):
- `feat:` - New feature
- `fix:` - Bug fix
- `docs:` - Documentation changes
- `test:` - Adding/updating tests
- `refactor:` - Code refactoring
- `perf:` - Performance improvements
- `chore:` - Maintenance tasks

### 5. Push and create PR
```bash
git push origin feature/your-feature-name
```

Then create a Pull Request on GitHub.

## Code Style Guide

### Python Style
- Follow [PEP 8](https://peps.python.org/pep-0008/)
- Use type hints for all function signatures
- Maximum line length: 100 characters
- Use `ruff` for formatting and linting

### Documentation
- Add docstrings to all public functions/classes
- Use Google-style docstrings
- Include examples for complex functionality

### Type Hints
```python
def convert_excel_to_ods(
    input_path: Path,
    output_path: Path,
    locale: str = "en-US",
) -> ConversionReport:
    """Convert Excel file to ODS format.

    Args:
        input_path: Path to input Excel file
        output_path: Path for output ODS file
        locale: Target locale (default: en-US)

    Returns:
        Conversion report with statistics and errors

    Raises:
        ConversionError: If conversion fails
    """
    ...
```

## Testing

### Running Tests
```bash
# All tests
pytest

# Unit tests only
pytest -m "not integration"

# Integration tests (requires LibreOffice)
pytest -m integration

# With coverage
pytest --cov=xlsliberator --cov-report=html
```

### Writing Tests
- Place unit tests in `tests/unit/`
- Place integration tests in `tests/it/`
- Use descriptive test names: `test_convert_excel_with_formulas_succeeds`
- Aim for >80% code coverage

### Test Structure
```python
def test_feature_name():
    """Test that feature does X when Y."""
    # Arrange
    input_data = setup_test_data()

    # Act
    result = function_under_test(input_data)

    # Assert
    assert result.success
    assert result.value == expected_value
```

## Pull Request Guidelines

### Before Submitting
- ✅ All tests pass
- ✅ Code is formatted (`ruff format`)
- ✅ No linting errors (`ruff check`)
- ✅ Type checking passes (`mypy src/`)
- ✅ Documentation is updated
- ✅ CHANGELOG.md is updated (for notable changes)

### PR Description Template
```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
How was this tested?

## Checklist
- [ ] Tests added/updated
- [ ] Documentation updated
- [ ] CHANGELOG updated
```

## Project Structure

```
xlsliberator/
├── src/xlsliberator/         # Source code
│   ├── api.py                # High-level API
│   ├── cli.py                # Command-line interface
│   ├── config.py             # Configuration management
│   ├── extract_vba.py        # VBA extraction
│   ├── vba2py_uno.py         # VBA→Python translation
│   ├── formula_*.py          # Formula handling
│   └── uno_conn.py           # LibreOffice UNO connection
├── tests/                    # Test suite
│   ├── unit/                 # Unit tests
│   ├── it/                   # Integration tests
│   └── data/                 # Test fixtures
├── rules/                    # YAML mapping rules
├── docs/                     # Documentation
└── .github/workflows/        # CI/CD workflows
```

## Issue Reporting

### Bug Reports
Include:
- XLSLiberator version
- Python version
- Operating system
- Minimal reproducible example
- Expected vs actual behavior
- Error messages/stack traces

### Feature Requests
Include:
- Use case description
- Proposed solution
- Alternatives considered
- Willingness to contribute implementation

## Getting Help

- **Issues**: [GitHub Issues](https://github.com/johannhartmann/xlsliberator/issues)
- **Discussions**: [GitHub Discussions](https://github.com/johannhartmann/xlsliberator/discussions)
- **Email**: johann-peter.hartmann@mayflower.de

## License

By contributing, you agree that your contributions will be licensed under the GPL-3.0-or-later license.
