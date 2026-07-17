# Repository Guidelines

## Project Structure & Module Organization

XLSLiberator is a Python package under `src/xlsliberator/`. Core entry points are `api.py`, `cli.py`, `mcp_server.py`, and `mcp_tools.py`; LibreOffice/UNO support lives in modules such as `uno_conn.py`, `lo_worker.py`, and `lo_worker_client.py`. Formula, validation, VBA translation, and macro embedding code use focused `formula_*`, `validation_*`, `vba*`, and `embed_macros.py` modules. Tests are split into `tests/unit/`, `tests/it/`, `tests/real/`, and `tests/bench/`. Mapping rules are in `rules/`, examples in `examples/`, and documentation in `docs/`.

## Build, Test, and Development Commands

Docker is the only supported development and runtime platform. The host shell may
run Docker, Git, and file operations only. Never start host Python, `uv`, PyUNO,
UNO, LibreOffice, or `soffice`, including for diagnostics. Common commands:

- `docker compose build test`: build the development image.
- `docker compose run --rm test ruff check src tests`: run lint checks.
- `docker compose run --rm test ruff format .`: format Python code.
- `docker compose run --rm test mypy src`: run static typing checks.
- `docker compose run --rm test pytest tests/unit`: run unit tests.
- `make test-integration`: run Docker-orchestrated LibreOffice integration tests.
- `make all`: run the Makefile CI-style quality sequence.

## Coding Style & Naming Conventions

Use Python 3.11+ with 4-space indentation and typed function signatures. Ruff enforces formatting and lint rules with a 100-character line target. Mypy is strict; prefer explicit return types and avoid `Any` unless an external UNO object forces it. Use `snake_case` for functions and modules, `PascalCase` for classes, and descriptive test names such as `test_convert_excel_with_formulas_succeeds`.

## Testing Guidelines

Pytest discovers `test_*.py`, `Test*` classes, and `test_*` functions. Unit tests belong in `tests/unit/`; LibreOffice or UNO integration tests belong in `tests/it/` and should use the `integration` marker. Skip cleanly when LibreOffice, pyuno, API keys, or real workbook fixtures are unavailable. Contributor docs target more than 80% coverage.

## Commit & Pull Request Guidelines

Use Conventional Commits, for example `feat: add conversion option`, `fix: handle UNO worker failure`, `test: cover macro execution`. Keep the first line concise and imperative. Do not mention AI assistance, Codex, Anthropic, or assistant tooling in commit messages.

Before opening a PR, run lint, type checks, and relevant tests. PR descriptions should summarize the change, list testing, link issues when applicable, and note documentation updates.

## Security & Configuration Tips

Do not commit secrets or local environment files. `ANTHROPIC_API_KEY` is optional
and only needed for VBA translation. LibreOffice is the sole target and is pinned
to full build `26.2.4.2`. LibreOffice, its bundled Python, UNO, and PyUNO run only
inside the repository's pinned office image. There is no host executable discovery,
host diagnostic, direct `soffice` fallback, or local PyUNO fallback.
