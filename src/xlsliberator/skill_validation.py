"""Strict validation for Deep Agents domain skills."""

from __future__ import annotations

import argparse
import re
from collections.abc import Sequence
from pathlib import Path
from typing import Any

import yaml

MAX_SKILL_BYTES = 128 * 1024
MAX_DESCRIPTION_LENGTH = 1024
MAX_COMPATIBILITY_LENGTH = 500

_SKILL_NAME = re.compile(r"^[a-z0-9]+(?:-[a-z0-9]+)*$")
_TOOL_NAME = re.compile(r"^[A-Za-z][A-Za-z0-9_.:-]{0,127}$")


class SkillValidationError(ValueError):
    """A domain skill violates the XLSLiberator metadata contract."""


def _frontmatter(text: str, path: Path) -> dict[str, Any]:
    if not text.startswith("---\n"):
        raise SkillValidationError(f"{path}: missing YAML frontmatter")
    marker = text.find("\n---\n", 4)
    if marker < 0:
        raise SkillValidationError(f"{path}: unterminated YAML frontmatter")
    try:
        metadata = yaml.safe_load(text[4:marker])
    except yaml.YAMLError as exc:
        raise SkillValidationError(f"{path}: invalid YAML frontmatter") from exc
    if not isinstance(metadata, dict):
        raise SkillValidationError(f"{path}: frontmatter must be a mapping")
    return metadata


def _required_string(metadata: dict[str, Any], key: str, path: Path) -> str:
    value = metadata.get(key)
    if not isinstance(value, str) or not value.strip():
        raise SkillValidationError(f"{path}: {key} must be a non-empty string")
    return value.strip()


def _tools(metadata: dict[str, Any], path: Path) -> list[str]:
    value = metadata.get("allowed-tools", metadata.get("recommended-tools"))
    if isinstance(value, str):
        tools = [tool for tool in re.split(r"[\s,]+", value) if tool]
    elif isinstance(value, list) and all(isinstance(tool, str) for tool in value):
        tools = [tool.strip() for tool in value if tool.strip()]
    else:
        raise SkillValidationError(
            f"{path}: allowed-tools or recommended-tools must declare tools"
        )
    if not tools or any(not _TOOL_NAME.fullmatch(tool) for tool in tools):
        raise SkillValidationError(f"{path}: tool declarations contain invalid names")
    return tools


def validate_skill(path: Path) -> dict[str, Any]:
    """Validate one SKILL.md and return its parsed metadata."""

    if path.is_symlink() or path.parent.is_symlink():
        raise SkillValidationError(f"{path}: symbolic links are forbidden")
    content = path.read_bytes()
    if len(content) > MAX_SKILL_BYTES:
        raise SkillValidationError(f"{path}: SKILL.md exceeds {MAX_SKILL_BYTES} bytes")
    try:
        text = content.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise SkillValidationError(f"{path}: SKILL.md must be UTF-8") from exc
    metadata = _frontmatter(text, path)

    name = _required_string(metadata, "name", path)
    if not _SKILL_NAME.fullmatch(name) or name != path.parent.name:
        raise SkillValidationError(
            f"{path}: name must match its lowercase-hyphen directory"
        )
    description = _required_string(metadata, "description", path)
    if (
        len(description) > MAX_DESCRIPTION_LENGTH
        or len(description.split()) < 6  # noqa: PLR2004
        or not re.search(r"\b(?:use|when)\b", description, flags=re.IGNORECASE)
    ):
        raise SkillValidationError(
            f"{path}: description must explain what the skill does and when to use it"
        )
    compatibility = _required_string(metadata, "compatibility", path)
    if len(compatibility) > MAX_COMPATIBILITY_LENGTH:
        raise SkillValidationError(
            f"{path}: compatibility exceeds {MAX_COMPATIBILITY_LENGTH} characters"
        )
    _tools(metadata, path)
    return metadata


def lint_skill_roots(roots: Sequence[Path]) -> list[str]:
    """Return deterministic errors across one or more skill roots."""

    errors: list[str] = []
    for root in roots:
        if not root.is_dir():
            errors.append(f"{root}: skill root is inaccessible")
            continue
        for path in sorted(root.glob("*/SKILL.md")):
            try:
                validate_skill(path)
            except (OSError, SkillValidationError) as exc:
                errors.append(str(exc))
    return errors


def main(argv: Sequence[str] | None = None) -> int:
    """Run the repository skill linter."""

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("roots", nargs="*", type=Path, default=[Path("skills")])
    args = parser.parse_args(argv)
    errors = lint_skill_roots(args.roots)
    for error in errors:
        print(error)
    if errors:
        print(f"skill lint failed with {len(errors)} error(s)")
        return 1
    print(f"skill lint passed for {len(args.roots)} root(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
