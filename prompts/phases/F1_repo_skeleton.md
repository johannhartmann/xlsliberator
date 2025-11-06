# Prompt F1 — Repo‑Skeleton & Tooling (OSS Setup)

**Ziel:** lauffähiges Grundgerüst mit QA‑Tools.

## Umgebung:
- **Python Environment:** conda environment `xlsliberator` (already activated)
- **Package Manager:** `uv` for fast dependency management
- **Tooling:** `ruff` (formatting & linting), `mypy` (type checking), `pytest` (testing)
- **Project Structure:** Modern Python module with `pyproject.toml`

## Aufgabe:

Erzeuge ein Repo `xlsliberator/` mit:

* `pyproject.toml` (Python 3.11+, modern build system with `uv`), deps: `openpyxl`, `odfpy`, `oletools`, `pyxlsb`, `pydantic`, `loguru`, `click`, `pytest`, `pytest-xdist`, `pytest-benchmark`, `mypy`, `ruff`.
* `src/xlsliberator/` (leere Module): `ir_models.py`, `extract_excel.py`, `extract_vba.py`, `formula_mapper.py`, `map_rules.py`, `uno_conn.py`, `write_ods.py`, `embed_macros.py`, `vba2py_uno.py`, `tables_reader.py`, `tables_to_uno.py`, `charts_reader.py`, `charts_to_uno.py`, `forms_parser.py`, `forms_to_uno.py`, `testing_lo.py`, `report.py`, `api.py`, `cli.py`.
* `rules/` (Platzhalter‑YAML): `formula_map.yaml`, `vba_api_map.yaml`, `event_map.yaml`, `forms_map.yaml`, `charts_map.yaml`.
* `tests/` + `tests/data/` (synthetische Mini‑Dateien als Fixtures).
* `README.md`, `LICENSE` (MIT), `ruff.toml`, `mypy.ini`, `.editorconfig`, `.gitignore`, `Makefile` (`fmt`, `lint`, `typecheck`, `test`).

## Tests & Metriken:
`pytest -q` ok; `ruff`/`mypy` laufen.

## Gate G1:
CI‑lokal grün (keine Tests fehlschlagen).

## Kommandos:
```bash
# Install dependencies with uv
uv pip install -e ".[dev]"

# Run quality checks
make fmt && make lint && make typecheck && make test
```

## Checklist Reference:
- [ ] Phase 0.1: Repo-Gerüst (src/, tests/, rules/, docs/) - `pytest`, `ruff`, `mypy` grün
