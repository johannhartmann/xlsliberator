# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- GitHub Actions CI/CD workflows for testing, linting, and security scanning
- Configuration management module (`config.py`) with environment variable support
- API key validation for Anthropic API
- Comprehensive CONTRIBUTING.md with development guidelines
- Pre-commit hooks configuration
- Project metadata updates (author, repository URLs)

### Changed
- Updated README with correct author information and repository links
- Improved package metadata in pyproject.toml

### Security
- Added dependency audit workflow
- Added secret scanning with Gitleaks
- Added SAST scanning with Bandit

## [0.1.0] - 2025-01-09

### Added
- Initial release
- Excel to LibreOffice Calc conversion using hybrid approach
- LibreOffice native conversion engine for 100% formula equivalence
- VBA-to-Python-UNO macro translation using LLM
- Deterministic AST-based formula transformation
- Named ranges fixing for native conversion
- Formula repair loop with retry logic
- Comprehensive test suite (99 tests)
- CLI interface with convert command
- Python API for programmatic use
- Conversion reporting (JSON and Markdown)
- Support for .xlsx, .xlsm, .xlsb, .xls formats
- GPL-3.0-or-later license

### Known Issues
- VBA translation requires Anthropic API key
- Some complex VBA patterns may require manual review
- Cross-workbook references need manual adjustment
- Integration tests may fail without proper LibreOffice setup

[Unreleased]: https://github.com/johannhartmann/xlsliberator/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/johannhartmann/xlsliberator/releases/tag/v0.1.0
