# Repository Guidelines

## Project Structure & Module Organization

XLSLiberator is a Python package under `src/xlsliberator/`. Core entry points are `api.py`, `cli.py`, `mcp_server.py`, and `mcp_tools.py`; LibreOffice/UNO support lives in modules such as `uno_conn.py`, `lo_worker.py`, and `lo_worker_client.py`. Formula, validation, VBA translation, and macro embedding code use focused `formula_*`, `validation_*`, `vba*`, and `embed_macros.py` modules. Tests are split into `tests/unit/`, `tests/it/`, `tests/real/`, and `tests/bench/`. Mapping rules are in `rules/`, examples in `examples/`, and documentation in `docs/`.

## Build, Test, and Development Commands

Install with `pip install -e ".[dev]"` or use `uv run` for tools. Common commands:

- `uv run ruff check src tests` or `make lint`: run lint checks.
- `uv run ruff format .` or `make fmt`: format Python code.
- `uv run mypy src` or `make typecheck`: run static typing checks.
- `uv run pytest` or `make test`: run the full test suite.
- `uv run pytest tests/unit`: run unit tests only.
- `uv run pytest -m integration`: run LibreOffice-dependent integration tests.
- `make all`: run the Makefile CI-style quality sequence.

## Coding Style & Naming Conventions

Use Python 3.11+ with 4-space indentation and typed function signatures. Ruff enforces formatting and lint rules with a 100-character line target. Mypy is strict; prefer explicit return types and avoid `Any` unless an external UNO object forces it. Use `snake_case` for functions and modules, `PascalCase` for classes, and descriptive test names such as `test_convert_excel_with_formulas_succeeds`.

## Testing Guidelines

Pytest discovers `test_*.py`, `Test*` classes, and `test_*` functions. Unit tests belong in `tests/unit/`; LibreOffice or UNO integration tests belong in `tests/it/` and should use the `integration` marker. Skip cleanly when LibreOffice, pyuno, API keys, or real workbook fixtures are unavailable. Contributor docs target more than 80% coverage.

## Commit & Pull Request Guidelines

Use Conventional Commits, for example `feat: add conversion option`, `fix: handle UNO worker failure`, `test: cover macro execution`. Keep the first line concise and imperative. Do not mention AI assistance, Codex, Anthropic, or assistant tooling in commit messages.

Before opening a PR, run lint, type checks, and relevant tests. PR descriptions should summarize the change, list testing, link issues when applicable, and note documentation updates.

## Security & Configuration Tips

Do not commit secrets or local environment files. `ANTHROPIC_API_KEY` is optional and only needed for VBA translation. LibreOffice 7.x with Python/UNO support is required for runtime validation and integration tests; simple conversion paths may still use `soffice` CLI fallback.
