# Excel → LibreOffice Calc Converter
## Implementation & Quality Gate Checklist (Feasibility One-Pager)

> Ziel: Schrittweise Machbarkeits-Validierung des `xlsliberator`-Prototyps (Excel-VBA → LibreOffice-Calc `.ods` mit eingebetteten Python-UNO-Makros)

---

### Phase 0 – Setup & Architektur
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **0.1** | Repo-Gerüst (src/, tests/, rules/, docs/) | `pytest`, `ruff`, `mypy` grün | ☑ |
| **0.2** | Feasibility-Plan + Scorecard | `docs/feasibility_plan.md` vorhanden | ☑ |
| **0.3** | LibreOffice-Headless-Harness (`uno_conn.py`) | 10/10 stabile Verbindungszyklen | ☑ |

---

### Phase 1 – Excel Ingestion & IR
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **1.1** | `extract_excel.py` – liest `.xlsx/.xlsm/.xlsb/.xls` | ≥ 99 % Formeln erkannt | ☑ |
| **1.2** | `ir_models.py` – WorkbookIR, SheetIR, NamedRangeIR | JSON-Serialisierung OK | ☑ |
| **1.3** | Performance-Benchmark (Ingestion) | ≥ 50 k Zellen/min | ☐ |

---

### Phase 2 – ODS Writer & Formula Mapping
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **2.1** | `formula_mapper.py` – 25 Kernfunktionen + Locale | ≥ 90 % syntaktisch korrekt | ☑ |
| **2.2** | `write_ods.py` – Werte + Formeln + NamedRanges | Recalc erfolgreich | ☑ |
| **2.3** | Formel-Vergleichstest (`testing_lo.py`) | ≥ 95 % Werte im Toleranzband (1e-9) | ☐ |

---

### Phase 3 – Makros & Events
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **3.1** | `extract_vba.py` – Module + Graph | 100 % Module erkannt | ☑ |
| **3.2** | `vba2py_uno.py` – Kern-Subset (Range/Cells/Worksheets) | Übersetzte Handlertests grün | ☑ |
| **3.3** | `embed_macros.py` – Scripts/python/*.py + Manifest | Event „on_open" feuert 1× | ☑ |

---

### Phase 4 – Erweiterungen: Tables, Charts, Forms
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **4.1** | `tables_reader.py` + `tables_to_uno.py` | ≥ 90 % Tabellen-Formeln korrekt | ☐ |
| **4.2** | `charts_reader.py` + `charts_to_uno.py` | ≥ 80 % Charts erzeugt (Serien+Titel) | ☐ |
| **4.3** | `forms_parser.py` + `forms_to_uno.py` | ≥ 70 % Controls & Events lauffähig | ☐ |

---

### Phase 5 – Integration, CLI & Reporting
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **5.1** | `api.py` + `cli.py` – End-to-End `convert()` | CLI-Smoke grün | ☑ |
| **5.2** | `report.py` – JSON/MD ConversionReport | Report enthält Metriken & Warnungen | ☑ |
| **5.3** | `tools/scorecard.py` – automatische Ampel | Scorecard generiert korrekt | ☐ |

---

### Phase 6 – Feasibility Validation (Real Data)
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **6.1** | Real-Dataset Test (z. B. Tippspiel-XLSM) | ≥ 1 Datei erfolgreich E2E | ☑ |
| **6.2** | Formel-Vergleich Real Data | ≥ 90 % Toleranzband erreicht | 🔄 |
| **6.3** | Performance Real Data | < 2 GB RAM / < 5 min / Datei | ☑ |

**✅ STRATEGIC DECISION MADE (2025-11-07):**

**Decision:** Use LibreOffice native conversion + VBA translation (Hybrid Approach)

**Rationale:**
- Manual translation achieved only 64% match rate
- LibreOffice native conversion expected to achieve 100% (same calc engine)
- Focus on unique value: VBA→Python-UNO translation

**Phase 6.2 - New Implementation Plan:**
| Sub-Phase | Task | Status |
|-----------|------|--------|
| **6.2.1** | Refactor `api.py` to use native conversion | ☐ |
| **6.2.2** | Test native conversion formula equivalence | ☐ |
| **6.2.3** | Integrate VBA translation with native ODS | ☐ |
| **6.2.4** | End-to-end validation | ☐ |
| **6.2.5** | Achieve ≥ 95% formula match rate | ☐ |

---

### Phase 7 – Stability, Security & Fallback
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **7.1** | Langläufer-Test (100 Open/Close Zyklen) | 100/100 stabil | ☐ |
| **7.2** | Sicherheitsprüfungen (VBA-Sandbox, Pfad-Härtung) | Keine Code-Ausführung, kein Leak | ☐ |
| **7.3** | Fallback-Import (auto SaveAs ODS) | Kein Hard-Fail bei Lücken | ☐ |

---

### Phase 8 – Windows Validator (Optional)
| Schritt | Deliverable | Gate / Metrik | Status |
|----------|--------------|---------------|---------|
| **8.1** | Excel-COM Vergleichstest | ≤ 1e-9 Abweichung | ☐ |
| **8.2** | Sandbox Execution | Keine UI/FS-Nebenwirkungen | ☐ |

---

### Phase 9 – Go/No-Go Criteria
| Kategorie | Grün | Gelb | Rot |
|------------|------|------|-----|
| Formel-Übersetzung | ≥ 90 % | 75–90 % | < 75 % |
| Formel-Gleichheit | ≥ 95 % | 85–95 % | < 85 % |
| Makro-Portierung | ≥ 80 % | 60–80 % | < 60 % |
| Ereignisse | ≥ 90 % | 70–90 % | < 70 % |
| Forms | ≥ 70 % | 50–70 % | < 50 % |
| Charts | ≥ 80 % | 60–80 % | < 60 % |
| Performance | Ziel erreicht | leicht drunter | stark drunter |
| Stabilität | ≥ 100/100 | 90–99 | < 90 |

---

### Hinweise
- **CI Integration:** jede Gate-Kategorie → Ampel in `feasibility_scorecard.md`
- **Feasibility erreicht**, wenn alle Gates **G1–G12** grün oder ≥ 80 % grün / keine „rot“
- **Sicherheits-Pflichten:** kein VBA-Runtime-Execute, nur statisch; COM-Tests isoliert
- **Reporting:** jede Pipeline-Ausgabe erzeugt aktualisierte Scorecard

---

✅ **Gesamtziel:**
Nach Abschluss aller grünen Gates produziert `xlsliberator convert *.xlsm → *.ods`
ein vollständig geöffnetes, berechenbares Calc-Dokument mit eingebetteten, funktional äquivalenten Python-Makros.

