# Prompt F5 — Formel‑Mapper v1 (Tokenizer + Locale)

**Ziel:** **sichere Syntax‑Übersetzung** (ohne Auswertung).

## Aufgabe:

Implementiere `formula_mapper.py` + `rules/formula_map.yaml`:

* Tokenizer (Funktionsnamen, Separatoren, Bezüge, Strings).
* Locale‑Aware: `,` vs `;` für `en-US`/`de-DE`.
* Mapping für ~25 Kernfunktionen; strukturierte Verweise (Stub) markieren.
* `tests/unit/test_formula_mapper.py`: Positiv/Negativ‑Fälle (u. a. `INDIRECT`, `OFFSET`, Fehlerwerte).

## Metrik:
≥ **90 %** syntaktisch korrekte Übersetzungen auf Test‑Korpus.

## Gate G5:
Bestehende `it`‑Tests weiter grün; Mapper‑Tests ≥ 90 % pass.

## Befehl:
```bash
pytest -q
```

## Checklist Reference:
- [ ] Phase 2.1: `formula_mapper.py` – 25 Kernfunktionen + Locale - ≥ 90 % syntaktisch korrekt
