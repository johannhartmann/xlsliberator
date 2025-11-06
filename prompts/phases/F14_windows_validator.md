# Prompt F14 — Windows‑Validator (optional)

**Ziel:** Excel‑COM‑Abgleich (Sandbox).

## Aufgabe:

Implementiere `tests/it/test_win_excel_validator.py` (skip unless Windows):

* Excel via `pywin32`: `CalculateFullRebuild`; lese Stichproben‑Zellen; vergleiche mit LO‑Werten.

## Metrik:
gleiche Werte (Toleranz 1e‑9) auf Stichprobe.

## Gate G14:
Test grün auf Windows; sonst übersprungen.

## Befehl:
```bash
pytest -q -k win_excel_validator
```

## Checklist Reference:
- [ ] Phase 8.1: Excel-COM Vergleichstest - ≤ 1e-9 Abweichung
- [ ] Phase 8.2: Sandbox Execution - Keine UI/FS-Nebenwirkungen
