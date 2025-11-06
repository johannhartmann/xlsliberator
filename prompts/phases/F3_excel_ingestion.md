# Prompt F3 — Excel‑Ingestion (alle Formate, ohne Auswertung)

**Ziel:** **WorkbookIR** aus `.xlsx/.xlsm/.xlsb/.xls` erzeugen.

## Aufgabe:

Implementiere `extract_excel.py` + `ir_models.py`:

* `.xlsx/.xlsm`: `openpyxl` (read‑only). Sammle Sheets, Werte‑Sample, **alle Formeln**, **NamedRanges**, Chart‑/Drawing‑Rels, ListObjects‑Hinweise, `vbaProject.bin` Flag.
* `.xlsb`: Werte + Formeln‑Stub via `pyxlsb` (markiere `formula_available=False`).
* `.xls`: Legacy‑Erkennung + Makro‑Flag (Details optional).
* `tests/unit/test_extract_excel.py`: synthetische `.xlsx/.xlsm` (ohne echte Makros) + Assertions auf IR‑Zählungen.

## Metrik:
#extrahierte Formeln / #Formeln ≥ **99 %** (synthetischer Fixpunkt).

## Gate G3:
IR enthält alle Formeln & NamedRanges der Testdateien; Laufzeit/Memory im Rahmen (notiere Peak).

## Befehl:
```bash
pytest tests/unit/test_extract_excel.py -q
```

## Checklist Reference:
- [ ] Phase 1.1: `extract_excel.py` – liest `.xlsx/.xlsm/.xlsb/.xls` - ≥ 99 % Formeln erkannt
- [ ] Phase 1.2: `ir_models.py` – WorkbookIR, SheetIR, NamedRangeIR - JSON-Serialisierung OK
- [ ] Phase 1.3: Performance-Benchmark (Ingestion) - ≥ 50 k Zellen/min
