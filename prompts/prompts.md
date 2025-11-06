Super – hier ist eine **inkrementelle Prompt‑Reihenfolge für „Claude Code“**, die den Converter **schrittweise** implementiert **und** in jeder Phase eine **frühe Machbarkeitsprüfung (Feasibility‑Gate)** einbaut.
Jeder Prompt enthält: **Ziel**, **Aufgaben/Artefakte**, **Tests & Metriken**, **Gate (Exit‑Kriterium)** und **Beispiel‑Kommandos**.
Kopiere die Prompts **nacheinander** in Claude Code; passe Platzhalter (`<...>`) an.

---

## Legende (kurz)

* **IR** = Intermediate Representation (ein neutrales Datenmodell für Workbook/Module/Forms/Charts).
* **Gate** = messbares Feasibility‑Kriterium; bei Nichterreichen → **Stop & Fix**, nicht weiterbauen.
* **Linux first**, Windows optional (Excel‑COM‑Validator).
* **Ziel:** `.ods` mit **eingebetteten Python‑UNO‑Makros** (Events, Forms), **Formel‑Parität** (funktional), **Tables** und **Charts**.

---

## Prompt F0 — Projekt‑Kickoff & Feasibility‑Roadmap

**Ziel:** Architektur + Gates festlegen.

**Prompt (Claude Code):**

> **Rolle:** Senior Python/UNO Engineer & Testarchitekt.
> **Aufgabe:** Erstelle eine **Roadmap** für `xlsliberator` (Excel→Calc‑Converter) mit **klaren Feasibility‑Gates**.
> **Lieferumfang:**
>
> 1. 10–15 Milestones mit Deliverables.
> 2. Pro Milestone: primäre Risiken, **Metriken**, **Gate** (quantitativ).
> 3. Minimal‑Scope v1 (Formeln/NamedRanges, Macro‑Stub, ODS‑Writer), v2 (VBA→Python Kernsubset + Events + Tables), v3 (Forms/Charts erweitert, Windows‑Validator).
> 4. Teststrategie (Unit/IT/E2E), Performanceziele, Locale (de-DE/en-US).
> 5. Artefaktbaum (`src/`, `tests/`, `rules/`, `docs/`).
>    **Output:** Markdown‑Plan `docs/feasibility_plan.md` + `docs/gates.md` (alle Gates tabellarisch).

---

## Prompt F1 — Repo‑Skeleton & Tooling (OSS Setup)

**Ziel:** lauffähiges Grundgerüst mit QA‑Tools.

**Prompt:**

> Erzeuge ein Repo `xlsliberator/` mit:
>
> * `pyproject.toml` (Python 3.11+), deps: `openpyxl`, `odfpy`, `oletools`, `pyxlsb`, `pydantic`, `loguru`, `click`, `pytest`, `pytest-xdist`, `pytest-benchmark`, `mypy`, `ruff`.
> * `src/xlsliberator/` (leere Module): `ir_models.py`, `extract_excel.py`, `extract_vba.py`, `formula_mapper.py`, `map_rules.py`, `uno_conn.py`, `write_ods.py`, `embed_macros.py`, `vba2py_uno.py`, `tables_reader.py`, `tables_to_uno.py`, `charts_reader.py`, `charts_to_uno.py`, `forms_parser.py`, `forms_to_uno.py`, `testing_lo.py`, `report.py`, `api.py`, `cli.py`.
> * `rules/` (Platzhalter‑YAML): `formula_map.yaml`, `vba_api_map.yaml`, `event_map.yaml`, `forms_map.yaml`, `charts_map.yaml`.
> * `tests/` + `tests/data/` (synthetische Mini‑Dateien als Fixtures).
> * `README.md`, `LICENSE` (MIT), `ruff.toml`, `mypy.ini`, `.editorconfig`, `.gitignore`, `Makefile` (`fmt`, `lint`, `typecheck`, `test`).
>   **Tests & Metriken:** `pytest -q` ok; `ruff`/`mypy` laufen.
>   **Gate G1:** CI‑lokal grün (keine Tests fehlschlagen).
>   **Kommandos:** `make fmt && make lint && make typecheck && make test`.

---

## Prompt F2 — LibreOffice‑Headless‑Harness (Connectivity)

**Ziel:** zuverlässige UNO‑Verbindung inkl. Guardrails.

**Prompt:**

> Implementiere `uno_conn.py`:
>
> * `connect_lo(host="127.0.0.1", port=2002, timeout=10) -> UnoCtx` (saubere Exceptions/Logs).
> * Helpers: `new_calc()`, `open_calc(url)`, `save_as_ods(doc, out_path)`, `recalc(doc)`, `get_sheet(doc, name|index)`, `get_cell(doc, addr)`.
> * `tests/it/test_uno_conn.py`: Überspringen, wenn LO nicht läuft (ENV `LO_SKIP_IT=1`).
>   **Feasibility‑Signal:** Stabiler Connect/Close über 10 Wiederholungen ohne Leak.
>   **Gate G2:** 10/10 Verbindungszyklen erfolgreich; `recalc()` auf leeres Doc ohne Fehler.
>   **Befehl:** `soffice --headless --accept=socket,host=127.0.0.1,port=2002;urp; & pytest tests/it/test_uno_conn.py -q`.

---

## Prompt F3 — Excel‑Ingestion (alle Formate, ohne Auswertung)

**Ziel:** **WorkbookIR** aus `.xlsx/.xlsm/.xlsb/.xls` erzeugen.

**Prompt:**

> Implementiere `extract_excel.py` + `ir_models.py`:
>
> * `.xlsx/.xlsm`: `openpyxl` (read‑only). Sammle Sheets, Werte‑Sample, **alle Formeln**, **NamedRanges**, Chart‑/Drawing‑Rels, ListObjects‑Hinweise, `vbaProject.bin` Flag.
> * `.xlsb`: Werte + Formeln‑Stub via `pyxlsb` (markiere `formula_available=False`).
> * `.xls`: Legacy‑Erkennung + Makro‑Flag (Details optional).
> * `tests/unit/test_extract_excel.py`: synthetische `.xlsx/.xlsm` (ohne echte Makros) + Assertions auf IR‑Zählungen.
>   **Metrik:** #extrahierte Formeln / #Formeln ≥ **99 %** (synthetischer Fixpunkt).
>   **Gate G3:** IR enthält alle Formeln & NamedRanges der Testdateien; Laufzeit/Memory im Rahmen (notiere Peak).
>   **Befehl:** `pytest tests/unit/test_extract_excel.py -q`.

---

## Prompt F4 — Mini‑ODS‑Writer (Werte + 10 Formeln)

**Ziel:** vom IR **zur ODS** (Minimalpfad), Recalc‑Smoke.

**Prompt:**

> Implementiere `write_ods.py`:
>
> * `build_calc_from_ir(ctx, ir, rules, locale="en-US") -> UnoDoc` (nur: Sheets, Werte, **10 häufige Excel‑Funktionen** als Formeln).
> * Nutze `formula_mapper.py` v0 (Hardcoded‑Map für: `IF`, `SUM`, `AVERAGE`, `SUMIF/SUMIFS`, `COUNTIF/COUNTIFS`, `INDEX/MATCH`, `VLOOKUP`, `XLOOKUP` wenn vorhanden).
> * `tests/it/test_ods_writer_smoke.py`: IR→ODS, `recalc`, 10 Zellen lesen und mit erwarteten Werten vergleichen.
>   **Metrik:** 10/10 Formeln korrekt in Calc.
>   **Gate G4:** Recalc liefert erwartete Werte (toleriert ±1e‑9).
>   **Befehl:** `pytest tests/it/test_ods_writer_smoke.py -q`.

---

## Prompt F5 — Formel‑Mapper v1 (Tokenizer + Locale)

**Ziel:** **sichere Syntax‑Übersetzung** (ohne Auswertung).

**Prompt:**

> Implementiere `formula_mapper.py` + `rules/formula_map.yaml`:
>
> * Tokenizer (Funktionsnamen, Separatoren, Bezüge, Strings).
> * Locale‑Aware: `,` vs `;` für `en-US`/`de-DE`.
> * Mapping für ~25 Kernfunktionen; strukturierte Verweise (Stub) markieren.
> * `tests/unit/test_formula_mapper.py`: Positiv/Negativ‑Fälle (u. a. `INDIRECT`, `OFFSET`, Fehlerwerte).
>   **Metrik:** ≥ **90 %** syntaktisch korrekte Übersetzungen auf Test‑Korpus.
>   **Gate G5:** Bestehende `it`‑Tests weiter grün; Mapper‑Tests ≥ 90 % pass.
>   **Befehl:** `pytest -q`.

---

## Prompt F6 — Macro‑Einbettung & Event‑Harness (Python‑UNO)

**Ziel:** Python‑Makros in `.ods` einbetten + Events verdrahten.

**Prompt:**

> Implementiere `embed_macros.py`:
>
> * `embed_python_macros(ods_path, py_modules)` → schreibe `Scripts/python/*.py` + `META-INF/manifest.xml` Updates.
> * Mini‑Modul `doc_events.py` mit `on_open(doc)` → schreibt Marker (z. B. `Sheet1.A1="OPEN_OK"`).
> * Registrierung: ODS so konfigurieren, dass `on_open` bei Dokument‑Öffnen läuft (per UNO Properties/Listeners).
> * `tests/it/test_macro_embed.py`: Öffnen `out.ods` headless, prüfen Markerzelle.
>   **Metrik:** Event feuert **genau einmal**.
>   **Gate G6:** Marker gesetzt, kein Crash/kein Mehrfachaufruf.
>   **Befehl:** `pytest tests/it/test_macro_embed.py -q`.

---

## Prompt F7 — VBA‑Extraktion (statisch) & Dependency‑Graph

**Ziel:** **VbaModuleIR** + Abhängigkeits‑Heatmap.

**Prompt:**

> Implementiere `extract_vba.py`:
>
> * `extract_vba_modules(path) -> list[VbaModuleIR]` (Standard/Klasse/Form + Quelltext), `build_vba_dependency_graph(mods)`.
> * Erkenne Tokens: `Range/Cells/Worksheets`, `Application.*`, `WorksheetFunction.*`, `UserForm`, `DoEvents`.
> * `tests/unit/test_extract_vba.py`: Golden‑Snippets (Strings/Dateien) → Modules & Graph.
>   **Metrik:** 100 % Module gezählt, Token‑Erkennung für Top‑APIs.
>   **Gate G7:** Graph baut fehlerfrei; Top‑APIs erkannt.
>   **Befehl:** `pytest tests/unit/test_extract_vba.py -q`.

---

## Prompt F8 — VBA→Python(UNO) Translator v1 (Subset)

**Ziel:** lauffähige **Minimal‑Portierung** + Event‑Hook.

**Prompt:**

> Implementiere `vba2py_uno.py`:
>
> * Unterstütze: `Sub/Function`, `Dim`, `If/ElseIf`, `For/Each`, `Select Case`, `With` (einfach), Aufrufe `Range/Cells/Worksheets`, einfache `WorksheetFunction` (SUM/COUNT/AVERAGE), `MsgBox`→Logger, `DoEvents`→NOOP.
> * `event_map.yaml`: `Workbook_Open`→`on_open`, `Worksheet_Change`→Sheet‑Listener.
> * `tests/unit/test_vba2py_uno.py`: VBA‑Snippet → erwarteter Python‑Code (Golden).
> * `tests/it/test_translated_macro_runs.py`: Übersetze Mini‑VBA, einbetten, ODS öffnen, Handler setzt Marker.
>   **Metrik:** übersetzte Snippets laufen; Marker korrekt.
>   **Gate G8:** IT‑Test grün.
>   **Befehl:** `pytest tests/it/test_translated_macro_runs.py -q`.

---

## Prompt F9 — Tables/ListObjects MVP

**Ziel:** Tabellen + strukturierte Verweise nutzbar machen.

**Prompt:**

> Implementiere `tables_reader.py` + `tables_to_uno.py`:
>
> * Lese ListObjects (Name, Header, DataRange, Spalten).
> * Erzeuge Calc‑DB‑Range + AutoFilter; benenne Bereiche.
> * `formula_mapper`: Regel für strukturierte Verweise → A1 (z. B. `=[@Amount]` → `A2` relativ zum Tabellenkontext).
> * `tests/it/test_tables_roundtrip.py`: Mini‑Excel mit Tabelle + Formel; ODS erzeugen, `recalc`, Werte prüfen.
>   **Metrik:** Tabellen‑Formeln korrekt für ≥ **90 %** der Fälle im Test.
>   **Gate G9:** IT‑Test grün.
>   **Befehl:** `pytest tests/it/test_tables_roundtrip.py -q`.

---

## Prompt F10 — Charts MVP (Line/Column)

**Ziel:** einfache Charts rekonstruieren.

**Prompt:**

> Implementiere `charts_reader.py` (lesen von `chart*.xml`, Serien, Kategorien) und `charts_to_uno.py` (Chart2‑Erzeugung).
>
> * `tests/it/test_charts_basic.py`: 1 Line + 1 Column Chart; verifiziere Serienanzahl, Titel, Legende (optional PNG‑Export).
>   **Metrik:** Serienanzahl = Original; Titel/Legende vorhanden.
>   **Gate G10:** IT‑Test grün.
>   **Befehl:** `pytest tests/it/test_charts_basic.py -q`.

---

## Prompt F11 — Formel‑Vergleich (funktional)

**Ziel:** **Wertgleichheit** im Calc nach Recalc.

**Prompt:**

> Implementiere in `testing_lo.py`:
>
> * `recalc_and_read(doc, addrs)`; `assert_almost_equal(expected, actual, tol=1e-9)`.
> * Sampling: pro Sheet n Formeln + alle „kritischen“ (`INDIRECT`, `OFFSET`).
> * `tests/it/test_formula_equivalence.py`: Erwarte Werte (aus Excel‑Cache oder Hand‑Referenz).
>   **Metrik:** ≥ **95 %** im Toleranzband.
>   **Gate G11:** Test grün; Ausreißer im Report dokumentiert.
>   **Befehl:** `pytest tests/it/test_formula_equivalence.py -q`.

---

## Prompt F12 — API/CLI + ConversionReport

**Ziel:** End‑to‑End‑Pipeline als **API** & **CLI**.

**Prompt:**

> Implementiere `api.py` + `cli.py` + `report.py`:
>
> * `convert(input_path, out_path, *, locale="de-DE", strict=False, sample=None, enable_charts=True, enable_forms=True) -> ConversionReport`.
> * Pipeline: `extract_excel` → `extract_vba` → `formula_mapper` → `write_ods` → `vba2py_uno` → `embed_macros` → `tables_to_uno` → `charts_to_uno` → `save`.
> * Report: #Zellen, #Formeln (übersetzt/unsupported), NamedRanges, Tables, Charts, Events, Makro‑APIs (gemappt/stubbed), Abweichungen, Laufzeiten.
> * `tests/it/test_cli_convert_smoke.py`: `xlsliberator convert in.xlsm out.ods`.
>   **Metrik:** Report erzeugt & plausibel; Exit‑Code 0.
>   **Gate G12:** CLI‑Smoke grün.
>   **Befehl:** `xlsliberator convert tests/data/sample.xlsm out/sample.ods --locale de-DE --strict`.

---

## Prompt F13 — Feasibility‑Scorecard (automatisch)

**Ziel:** Gates messbar zusammenfassen.

**Prompt:**

> Erzeuge `tools/scorecard.py`: liest `ConversionReport` + Benchmarks und schreibt `feasibility_scorecard.md` (Ampel pro Domäne).
>
> * `tests/unit/test_scorecard.py`: Snapshot‑Test.
>   **Metrik:** Scorecard zeigt G4–G12 Status.
>   **Gate G13:** Scorecard generiert; alle bisher grünen Gates korrekt gespiegelt.
>   **Befehl:** `python -m tools.scorecard out/report.json > out/feasibility_scorecard.md`.

---

## Prompt F14 — Windows‑Validator (optional)

**Ziel:** Excel‑COM‑Abgleich (Sandbox).

**Prompt:**

> Implementiere `tests/it/test_win_excel_validator.py` (skip unless Windows):
>
> * Excel via `pywin32`: `CalculateFullRebuild`; lese Stichproben‑Zellen; vergleiche mit LO‑Werten.
>   **Metrik:** gleiche Werte (Toleranz 1e‑9) auf Stichprobe.
>   **Gate G14:** Test grün auf Windows; sonst übersprungen.
>   **Befehl:** `pytest -q -k win_excel_validator`.

---

## Prompt F15 — Performance‑Mikrobenchmarks & Stabilität

**Ziel:** Durchsatz/Peak‑RAM/Stabilität prüfen.

**Prompt:**

> Ergänze Benchmarks (`pytest-benchmark`) für:
>
> * Ingestion (Zellen/s), Formula‑Mapping (Formeln/s), ODS‑Schreiben (Zellen/s).
> * Langläufer: 100 LO‑Open/Close‑Zyklen headless, Memory‑Peak loggen.
>   **Metrik:** Zielbereiche dokumentiert; keine Crashs/Leaks.
>   **Gate G15:** Benchmarks laufen, Stabilität 100/100.
>   **Befehl:** `pytest tests/bench -q` und `pytest tests/it/test_lo_stability.py -q`.

---

## Prompt F16 — Real‑Dataset‑Probe (Tippspiel)

**Ziel:** Frühe Realitätsprobe auf deinem Korpus.

**Prompt:**

> Lege `tests/real/datasets.yaml` mit Pfaden zu realen `.xlsm` an (z. B. Tipp‑Spiel).
>
> * `tests/real/test_convert_real.py`: iteriere Dateien → `convert` → prüfe:
>
>   1. ODS existiert & öffnet,
>   2. `recalc`,
>   3. 10 zufällige Formelzellen im Toleranzband,
>   4. Macro‑Open‑Marker gesetzt,
>   5. Report speichert Unsupported‑Listen.
>      **Metrik:** ≥ **80 %** der Tests grün beim ersten Lauf; Report zeigt Lücken gezielt.
>      **Gate G16:** Mindestens 1 Datei erfolgreich E2E; Scorecard aktualisiert.
>      **Befehl:** `pytest tests/real/test_convert_real.py -q`.

---

## Prompt F17 — Fallback‑Pfad (Full‑Auto‑Sicherung)

**Ziel:** „Immer Ergebnis“ auch bei Lücken.

**Prompt:**

> Implementiere optionalen Fallback in `api.py`:
>
> * Wenn bestimmter Anteil der Formeln/Charts/Forms „unsupported“ → **Plan B**: öffne Original in LO, `storeAsURL` `.ods` (Programmatic Import), dann **Nachbearbeitung** (NamedRanges/Events/Makros einbetten).
> * Report markiert Fallback‑Nutzung pro Feature.
>   **Metrik:** Kein Hard‑Fail; Ergebnis‑ODS immer erzeugt (sofern LO importiert).
>   **Gate G17:** Fallback greift, wenn Mapper‑Coverage < Schwellwert, und E2E bleibt grün.
>   **Befehl:** `xlsliberator convert in.xlsm out.ods --allow-fallback`.

---

# Hinweise zur Anwendung

* Führe **jeden Prompt** in Claude Code aus; committe die Artefakte.
* **Stoppe** bei einem roten **Gate**, fixiere gezielt und **wiederhole** den Prompt.
* Halte **Locale** konsistent (für de‑DE ggf. Semikolon‑Trenner in `formula_mapper`).
* Für **Makro‑Tests** nutze einfache Handler (Markerzellen) zur eindeutigen Verifikation.
* **Security:** VBA nur **statisch** analysieren; Excel‑COM‑Validator nur in isolierter Windows‑Umgebung.

---

Wenn du möchtest, gebe ich dir zusätzlich eine **„One‑Pager Checkliste“** (Markdown‑Vorlage) für die Gates (G1–G17), die ihr im Repo unter `docs/gates.md` pflegt und die CI nach jedem Lauf automatisch aktualisiert.

