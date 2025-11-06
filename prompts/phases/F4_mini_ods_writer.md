# Prompt F4 — Mini‑ODS‑Writer (Werte + 10 Formeln)

**Ziel:** vom IR **zur ODS** (Minimalpfad), Recalc‑Smoke.

## Aufgabe:

Implementiere `write_ods.py`:

* `build_calc_from_ir(ctx, ir, rules, locale="en-US") -> UnoDoc` (nur: Sheets, Werte, **10 häufige Excel‑Funktionen** als Formeln).
* Nutze `formula_mapper.py` v0 (Hardcoded‑Map für: `IF`, `SUM`, `AVERAGE`, `SUMIF/SUMIFS`, `COUNTIF/COUNTIFS`, `INDEX/MATCH`, `VLOOKUP`, `XLOOKUP` wenn vorhanden).
* `tests/it/test_ods_writer_smoke.py`: IR→ODS, `recalc`, 10 Zellen lesen und mit erwarteten Werten vergleichen.

## Metrik:
10/10 Formeln korrekt in Calc.

## Gate G4:
Recalc liefert erwartete Werte (toleriert ±1e‑9).

## Befehl:
```bash
pytest tests/it/test_ods_writer_smoke.py -q
```

## Checklist Reference:
- [ ] Phase 2.2: `write_ods.py` – Werte + Formeln + NamedRanges - Recalc erfolgreich
