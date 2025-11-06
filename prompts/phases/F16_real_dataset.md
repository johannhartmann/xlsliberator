# Prompt F16 — Real‑Dataset‑Probe (Tippspiel)

**Ziel:** Frühe Realitätsprobe auf deinem Korpus.

## Aufgabe:

Lege `tests/real/datasets.yaml` mit Pfaden zu realen `.xlsm` an (z. B. Tipp‑Spiel).

* `tests/real/test_convert_real.py`: iteriere Dateien → `convert` → prüfe:

  1. ODS existiert & öffnet,
  2. `recalc`,
  3. 10 zufällige Formelzellen im Toleranzband,
  4. Macro‑Open‑Marker gesetzt,
  5. Report speichert Unsupported‑Listen.

## Metrik:
≥ **80 %** der Tests grün beim ersten Lauf; Report zeigt Lücken gezielt.

## Gate G16:
Mindestens 1 Datei erfolgreich E2E; Scorecard aktualisiert.

## Befehl:
```bash
pytest tests/real/test_convert_real.py -q
```

## Checklist Reference:
- [ ] Phase 6.1: Real-Dataset Test (z. B. Tippspiel-XLSM) - ≥ 1 Datei erfolgreich E2E
- [ ] Phase 6.2: Formel-Vergleich Real Data - ≥ 90 % Toleranzband erreicht
