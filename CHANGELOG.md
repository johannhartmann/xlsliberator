# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Embedded pinned upstream Open-SWE runtime and XLSLiberator migration graph
- Authenticated Open-SWE web transport with fail-closed dependency readiness
- Pinned LibreOffice `26.2.4.2` application and source-build images
- GitHub Actions CI/CD workflows for testing, linting, and security scanning
- Configuration management module (`config.py`) with environment variable support

### Changed
- Docker is now the only supported application, test, and LibreOffice runtime platform
- Open-SWE is now the only supported agent and migration orchestrator
- Removed the embedded provider-backed translator and its compatibility switches
- Removed the unused corpus, demo, repair-promotion, skill-registry, and generated
  capability-report subsystems
- Simplified CI and release workflows around the supported Docker stack

### Security
- Pinned `setuptools` to 83.0.0 in application and audit images to address
  PYSEC-2026-3447
- Added dependency audit workflow
- Added secret scanning with Gitleaks
- Added SAST scanning with Bandit

## [0.1.0] - 2025-01-09

### Added
- Initial release
- Excel to LibreOffice Calc conversion using hybrid approach
- LibreOffice native conversion engine; formula equivalence requires scenario evidence
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
- VBA translation in this historical release required an external model API key
- Some complex VBA patterns may require manual review
- Cross-workbook references need manual adjustment
- Integration tests may fail without proper LibreOffice setup

[Unreleased]: https://github.com/johannhartmann/xlsliberator/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/johannhartmann/xlsliberator/releases/tag/v0.1.0
