from __future__ import annotations

from pathlib import Path

import pytest

from xlsliberator.skill_validation import (
    MAX_SKILL_BYTES,
    SkillValidationError,
    lint_skill_roots,
    validate_skill,
)


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
