# Prompt F2 — LibreOffice‑Headless‑Harness (Connectivity)

**Ziel:** zuverlässige UNO‑Verbindung inkl. Guardrails.

## Aufgabe:

Implementiere `uno_conn.py`:

* `connect_lo(host="127.0.0.1", port=2002, timeout=10) -> UnoCtx` (saubere Exceptions/Logs).
* Helpers: `new_calc()`, `open_calc(url)`, `save_as_ods(doc, out_path)`, `recalc(doc)`, `get_sheet(doc, name|index)`, `get_cell(doc, addr)`.
* `tests/it/test_uno_conn.py`: Überspringen, wenn LO nicht läuft (ENV `LO_SKIP_IT=1`).

## Feasibility‑Signal:
Stabiler Connect/Close über 10 Wiederholungen ohne Leak.

## Gate G2:
10/10 Verbindungszyklen erfolgreich; `recalc()` auf leeres Doc ohne Fehler.

## Befehl:
```bash
soffice --headless --accept="socket,host=127.0.0.1,port=2002;urp;" &
pytest tests/it/test_uno_conn.py -q
```

## Checklist Reference:
- [ ] Phase 0.3: LibreOffice-Headless-Harness (`uno_conn.py`) - 10/10 stabile Verbindungszyklen
