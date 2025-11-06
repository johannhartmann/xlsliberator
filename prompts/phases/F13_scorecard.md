# Prompt F13 — Feasibility‑Scorecard (automatisch)

**Ziel:** Gates messbar zusammenfassen.

## Aufgabe:

Erzeuge `tools/scorecard.py`: liest `ConversionReport` + Benchmarks und schreibt `feasibility_scorecard.md` (Ampel pro Domäne).

* `tests/unit/test_scorecard.py`: Snapshot‑Test.

## Metrik:
Scorecard zeigt G4–G12 Status.

## Gate G13:
Scorecard generiert; alle bisher grünen Gates korrekt gespiegelt.

## Befehl:
```bash
python -m tools.scorecard out/report.json > out/feasibility_scorecard.md
```

## Checklist Reference:
- [ ] Phase 5.3: `tools/scorecard.py` – automatische Ampel - Scorecard generiert korrekt
