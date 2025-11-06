# Prompt F6 — Macro‑Einbettung & Event‑Harness (Python‑UNO)

**Ziel:** Python‑Makros in `.ods` einbetten + Events verdrahten.

## Aufgabe:

Implementiere `embed_macros.py`:

* `embed_python_macros(ods_path, py_modules)` → schreibe `Scripts/python/*.py` + `META-INF/manifest.xml` Updates.
* Mini‑Modul `doc_events.py` mit `on_open(doc)` → schreibt Marker (z. B. `Sheet1.A1="OPEN_OK"`).
* Registrierung: ODS so konfigurieren, dass `on_open` bei Dokument‑Öffnen läuft (per UNO Properties/Listeners).
* `tests/it/test_macro_embed.py`: Öffnen `out.ods` headless, prüfen Markerzelle.

## Metrik:
Event feuert **genau einmal**.

## Gate G6:
Marker gesetzt, kein Crash/kein Mehrfachaufruf.

## Befehl:
```bash
pytest tests/it/test_macro_embed.py -q
```

## Checklist Reference:
- [ ] Phase 3.3: `embed_macros.py` – Scripts/python/*.py + Manifest - Event „on_open" feuert 1×
