# Prompt F17 — Fallback‑Pfad (Full‑Auto‑Sicherung)

**Ziel:** „Immer Ergebnis" auch bei Lücken.

## Aufgabe:

Implementiere optionalen Fallback in `api.py`:

* Wenn bestimmter Anteil der Formeln/Charts/Forms „unsupported" → **Plan B**: öffne Original in LO, `storeAsURL` `.ods` (Programmatic Import), dann **Nachbearbeitung** (NamedRanges/Events/Makros einbetten).
* Report markiert Fallback‑Nutzung pro Feature.

## Metrik:
Kein Hard‑Fail; Ergebnis‑ODS immer erzeugt (sofern LO importiert).

## Gate G17:
Fallback greift, wenn Mapper‑Coverage < Schwellwert, und E2E bleibt grün.

## Befehl:
```bash
xlsliberator convert in.xlsm out.ods --allow-fallback
```

## Checklist Reference:
- [ ] Phase 7.3: Fallback-Import (auto SaveAs ODS) - Kein Hard-Fail bei Lücken
