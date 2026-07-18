from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from xlsliberator.skill_validation import (
    MAX_SKILL_BYTES,
    SkillValidationError,
    lint_skill_roots,
    validate_skill,
)

PROJECT_ROOT = Path(__file__).parents[2]
SKILL_ROOT = PROJECT_ROOT / "skills"
DISCOVERY_FIXTURE = PROJECT_ROOT / "tests" / "fixtures" / "skill_discovery.yaml"
PROMPT_11_SKILLS = {
    "workbook-forensics",
    "migration-planning",
    "migration-test-design",
    "migration-mutation-testing",
    "ods-package-surgery",
    "workbook-failure-minimization",
    "secure-workbook-execution",
    "visual-validation",
}
PROMPT_12_SKILLS = {
    "vba-to-python-uno",
    "formula-migration",
    "userform-to-uno",
    "activex-to-open-controls",
    "windows-dependency-replacement",
    "libreoffice-debugging",
    "libreoffice-core-patching",
    "open-service-adapter",
}


def _write_skill(root: Path, name: str, frontmatter: str | None = None) -> Path:
    directory = root / name
    directory.mkdir()
    content = frontmatter or (
        "---\n"
        f"name: {name}\n"
        "description: Use this skill when workbook migration evidence needs deterministic checks.\n"
        "compatibility: Docker-only XLSLiberator with LibreOffice 26.2.4.2.\n"
        "recommended-tools: xlsprobe migration-check\n"
        "---\n\n"
        "# Workflow\n"
    )
    path = directory / "SKILL.md"
    path.write_text(content, encoding="utf-8")
    return path


def test_validate_skill_accepts_required_metadata(tmp_path: Path) -> None:
    path = _write_skill(tmp_path, "workbook-forensics")

    metadata = validate_skill(path)

    assert metadata["name"] == "workbook-forensics"


@pytest.mark.parametrize(
    ("name", "frontmatter", "message"),
    [
        (
            "workbook-forensics",
            "# missing frontmatter\n",
            "missing YAML frontmatter",
        ),
        (
            "workbook-forensics",
            "---\nname: [broken\n---\n",
            "invalid YAML frontmatter",
        ),
        (
            "workbook-forensics",
            (
                "---\n"
                "name: wrong-name\n"
                "description: Use this skill when workbook migration checks need guidance.\n"
                "compatibility: Docker-only XLSLiberator.\n"
                "allowed-tools: xlsprobe\n"
                "---\n"
            ),
            "name must match",
        ),
    ],
)
def test_validate_skill_rejects_invalid_metadata(
    tmp_path: Path,
    name: str,
    frontmatter: str,
    message: str,
) -> None:
    path = _write_skill(tmp_path, name, frontmatter)

    with pytest.raises(SkillValidationError, match=message):
        validate_skill(path)


def test_validate_skill_rejects_oversized_skill(tmp_path: Path) -> None:
    path = _write_skill(tmp_path, "workbook-forensics")
    path.write_bytes(path.read_bytes() + b"x" * MAX_SKILL_BYTES)

    with pytest.raises(SkillValidationError, match="exceeds"):
        validate_skill(path)


def test_lint_skill_roots_reports_inaccessible_source(tmp_path: Path) -> None:
    errors = lint_skill_roots([tmp_path / "missing"])

    assert errors == [f"{tmp_path / 'missing'}: skill root is inaccessible"]


def _words(text: str) -> set[str]:
    return {
        word.strip(".,:;()[]`").lower()
        for word in text.split()
        if len(word.strip(".,:;()[]`")) >= 4
    }


def test_project_skill_catalog_is_complete_and_valid() -> None:
    skill_paths = sorted(SKILL_ROOT.glob("*/SKILL.md"))

    assert {path.parent.name for path in skill_paths} == PROMPT_11_SKILLS | PROMPT_12_SKILLS
    assert lint_skill_roots([SKILL_ROOT]) == []


@pytest.mark.parametrize("skill_name", sorted(PROMPT_11_SKILLS))
def test_forensics_and_planning_skills_have_operational_contract(skill_name: str) -> None:
    text = (SKILL_ROOT / skill_name / "SKILL.md").read_text(encoding="utf-8")

    for section in (
        "## Use when",
        "## Inputs and outputs",
        "## Tool sequence",
        "## Failure handling",
        "## Acceptance checklist",
        "## Examples",
    ):
        assert section in text
    assert "Adversarial:" in text


@pytest.mark.parametrize("skill_name", sorted(PROMPT_12_SKILLS))
def test_specialist_skills_have_tested_examples_and_global_anti_patterns(
    skill_name: str,
) -> None:
    text = (SKILL_ROOT / skill_name / "SKILL.md").read_text(encoding="utf-8")

    assert "## Tested examples" in text
    assert "## Global anti-patterns" in text
    assert "No Excel worker" in text
    assert "VBA runtime" in text
    assert "`ExcelContext` expansion" in text
    assert "provider-specific core" in text
    assert "success without target execution" in text


def test_discovery_fixture_selects_unique_skill_from_descriptions() -> None:
    fixtures = yaml.safe_load(DISCOVERY_FIXTURE.read_text(encoding="utf-8"))
    descriptions = {
        path.parent.name: str(validate_skill(path)["description"])
        for path in sorted(SKILL_ROOT.glob("*/SKILL.md"))
    }

    for fixture in fixtures:
        query_words = _words(fixture["query"])
        scores = {
            name: len(query_words & _words(description))
            for name, description in descriptions.items()
        }
        best_score = max(scores.values())
        winners = {name for name, score in scores.items() if score == best_score}
        assert winners == {fixture["expected"]}, fixture["query"]
