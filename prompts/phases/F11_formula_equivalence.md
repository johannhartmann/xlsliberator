# Prompt F11 — Formel‑Vergleich (funktional)

**Ziel:** **Wertgleichheit** im Calc nach Recalc.

## Aufgabe:

Implementiere in `testing_lo.py`:

* `recalc_and_read(doc, addrs)`; `assert_almost_equal(expected, actual, tol=1e-9)`.
* Sampling: pro Sheet n Formeln + alle „kritischen" (`INDIRECT`, `OFFSET`).
* `tests/it/test_formula_equivalence.py`: Erwarte Werte (aus Excel‑Cache oder Hand‑Referenz).

## Metrik:
≥ **95 %** im Toleranzband.

## Gate G11:
Test grün; Ausreißer im Report dokumentiert.

## Befehl:
```bash
pytest tests/it/test_formula_equivalence.py -q
```

## Checklist Reference:
- [ ] Phase 2.3: Formel-Vergleichstest (`testing_lo.py`) - ≥ 95 % Werte im Toleranzband (1e-9)
