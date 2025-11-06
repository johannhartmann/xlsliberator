# Prompt F0 — Projekt‑Kickoff & Feasibility‑Roadmap

**Ziel:** Architektur + Gates festlegen.

**Rolle:** Senior Python/UNO Engineer & Testarchitekt.

## Umgebung:
- **Python Environment:** conda environment `xlsliberator` (already activated)
- **Package Manager:** `uv` for fast dependency management
- **Tooling:** `ruff` (formatting & linting), `mypy` (type checking), `pytest` (testing)
- **Project Structure:** Modern Python module with `pyproject.toml`

**Aufgabe:** Erstelle eine **Roadmap** für `xlsliberator` (Excel→Calc‑Converter) mit **klaren Feasibility‑Gates**.

## Lieferumfang:

1. 10–15 Milestones mit Deliverables.
2. Pro Milestone: primäre Risiken, **Metriken**, **Gate** (quantitativ).
3. Minimal‑Scope v1 (Formeln/NamedRanges, Macro‑Stub, ODS‑Writer), v2 (VBA→Python Kernsubset + Events + Tables), v3 (Forms/Charts erweitert, Windows‑Validator).
4. Teststrategie (Unit/IT/E2E), Performanceziele, Locale (de-DE/en-US).
5. Artefaktbaum (`src/`, `tests/`, `rules/`, `docs/`).

**Output:** Markdown‑Plan `docs/feasibility_plan.md` + `docs/gates.md` (alle Gates tabellarisch).

## Checklist Reference:
- [ ] Phase 0.1: Repo-Gerüst (src/, tests/, rules/, docs/)
- [ ] Phase 0.2: Feasibility-Plan + Scorecard
