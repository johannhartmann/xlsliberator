# Prompt F9 — Tables/ListObjects MVP

**Ziel:** Tabellen + strukturierte Verweise nutzbar machen.

## Aufgabe:

Implementiere `tables_reader.py` + `tables_to_uno.py`:

* Lese ListObjects (Name, Header, DataRange, Spalten).
* Erzeuge Calc‑DB‑Range + AutoFilter; benenne Bereiche.
* `formula_mapper`: Regel für strukturierte Verweise → A1 (z. B. `=[@Amount]` → `A2` relativ zum Tabellenkontext).
* `tests/it/test_tables_roundtrip.py`: Mini‑Excel mit Tabelle + Formel; ODS erzeugen, `recalc`, Werte prüfen.

## Metrik:
Tabellen‑Formeln korrekt für ≥ **90 %** der Fälle im Test.

## Gate G9:
IT‑Test grün.

## Befehl:
```bash
pytest tests/it/test_tables_roundtrip.py -q
```

## Checklist Reference:
- [ ] Phase 4.1: `tables_reader.py` + `tables_to_uno.py` - ≥ 90 % Tabellen-Formeln korrekt
