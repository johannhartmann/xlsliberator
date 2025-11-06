# Prompt F15 — Performance‑Mikrobenchmarks & Stabilität

**Ziel:** Durchsatz/Peak‑RAM/Stabilität prüfen.

## Aufgabe:

Ergänze Benchmarks (`pytest-benchmark`) für:

* Ingestion (Zellen/s), Formula‑Mapping (Formeln/s), ODS‑Schreiben (Zellen/s).
* Langläufer: 100 LO‑Open/Close‑Zyklen headless, Memory‑Peak loggen.

## Metrik:
Zielbereiche dokumentiert; keine Crashs/Leaks.

## Gate G15:
Benchmarks laufen, Stabilität 100/100.

## Befehl:
```bash
pytest tests/bench -q
pytest tests/it/test_lo_stability.py -q
```

## Checklist Reference:
- [ ] Phase 7.1: Langläufer-Test (100 Open/Close Zyklen) - 100/100 stabil
- [ ] Phase 6.3: Performance Real Data - < 2 GB RAM / < 5 min / Datei
