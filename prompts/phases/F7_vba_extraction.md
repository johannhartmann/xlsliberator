# Prompt F7 — VBA‑Extraktion (statisch) & Dependency‑Graph

**Ziel:** **VbaModuleIR** + Abhängigkeits‑Heatmap.

## Aufgabe:

Implementiere `extract_vba.py`:

* `extract_vba_modules(path) -> list[VbaModuleIR]` (Standard/Klasse/Form + Quelltext), `build_vba_dependency_graph(mods)`.
* Erkenne Tokens: `Range/Cells/Worksheets`, `Application.*`, `WorksheetFunction.*`, `UserForm`, `DoEvents`.
* `tests/unit/test_extract_vba.py`: Golden‑Snippets (Strings/Dateien) → Modules & Graph.

## Metrik:
100 % Module gezählt, Token‑Erkennung für Top‑APIs.

## Gate G7:
Graph baut fehlerfrei; Top‑APIs erkannt.

## Befehl:
```bash
pytest tests/unit/test_extract_vba.py -q
```

## Checklist Reference:
- [ ] Phase 3.1: `extract_vba.py` – Module + Graph - 100 % Module erkannt
