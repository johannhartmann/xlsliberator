# Prompt F12 — API/CLI + ConversionReport

**Ziel:** End‑to‑End‑Pipeline als **API** & **CLI**.

## Aufgabe:

Implementiere `api.py` + `cli.py` + `report.py`:

* `convert(input_path, out_path, *, locale="de-DE", strict=False, sample=None, enable_charts=True, enable_forms=True) -> ConversionReport`.
* Pipeline: `extract_excel` → `extract_vba` → `formula_mapper` → `write_ods` → `vba2py_uno` → `embed_macros` → `tables_to_uno` → `charts_to_uno` → `save`.
* Report: #Zellen, #Formeln (übersetzt/unsupported), NamedRanges, Tables, Charts, Events, Makro‑APIs (gemappt/stubbed), Abweichungen, Laufzeiten.
* `tests/it/test_cli_convert_smoke.py`: `xlsliberator convert in.xlsm out.ods`.

## Metrik:
Report erzeugt & plausibel; Exit‑Code 0.

## Gate G12:
CLI‑Smoke grün.

## Befehl:
```bash
xlsliberator convert tests/data/sample.xlsm out/sample.ods --locale de-DE --strict
```

## Checklist Reference:
- [ ] Phase 5.1: `api.py` + `cli.py` – End-to-End `convert()` - CLI-Smoke grün
- [ ] Phase 5.2: `report.py` – JSON/MD ConversionReport - Report enthält Metriken & Warnungen
