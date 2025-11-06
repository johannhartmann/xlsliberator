# Prompt F10 — Charts MVP (Line/Column)

**Ziel:** einfache Charts rekonstruieren.

## Aufgabe:

Implementiere `charts_reader.py` (lesen von `chart*.xml`, Serien, Kategorien) und `charts_to_uno.py` (Chart2‑Erzeugung).

* `tests/it/test_charts_basic.py`: 1 Line + 1 Column Chart; verifiziere Serienanzahl, Titel, Legende (optional PNG‑Export).

## Metrik:
Serienanzahl = Original; Titel/Legende vorhanden.

## Gate G10:
IT‑Test grün.

## Befehl:
```bash
pytest tests/it/test_charts_basic.py -q
```

## Checklist Reference:
- [ ] Phase 4.2: `charts_reader.py` + `charts_to_uno.py` - ≥ 80 % Charts erzeugt (Serien+Titel)
