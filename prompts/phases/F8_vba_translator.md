# Prompt F8 — VBA→Python(UNO) Translator v1 (Subset)

**Ziel:** lauffähige **Minimal‑Portierung** + Event‑Hook.

## Aufgabe:

Implementiere `vba2py_uno.py`:

* Unterstütze: `Sub/Function`, `Dim`, `If/ElseIf`, `For/Each`, `Select Case`, `With` (einfach), Aufrufe `Range/Cells/Worksheets`, einfache `WorksheetFunction` (SUM/COUNT/AVERAGE), `MsgBox`→Logger, `DoEvents`→NOOP.
* `event_map.yaml`: `Workbook_Open`→`on_open`, `Worksheet_Change`→Sheet‑Listener.
* `tests/unit/test_vba2py_uno.py`: VBA‑Snippet → erwarteter Python‑Code (Golden).
* `tests/it/test_translated_macro_runs.py`: Übersetze Mini‑VBA, einbetten, ODS öffnen, Handler setzt Marker.

## Metrik:
übersetzte Snippets laufen; Marker korrekt.

## Gate G8:
IT‑Test grün.

## Befehl:
```bash
pytest tests/it/test_translated_macro_runs.py -q
```

## Checklist Reference:
- [ ] Phase 3.2: `vba2py_uno.py` – Kern-Subset (Range/Cells/Worksheets) - Übersetzte Handlertests grün
