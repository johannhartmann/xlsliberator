"""Architecture guards for the deterministic XLSLiberator toolbelt."""

from __future__ import annotations

import ast
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PACKAGE = ROOT / "src" / "xlsliberator"
PROVIDER_ROOTS = {"anthropic", "openai", "langchain", "open_swe"}


def _imports(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    imported: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(alias.name.split(".", 1)[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".", 1)[0])
    return imported


def test_core_has_no_provider_or_agent_framework_imports() -> None:
    violations: dict[str, list[str]] = {}
    for path in PACKAGE.rglob("*.py"):
        if "legacy_agent" in path.parts:
            continue
        forbidden = sorted(_imports(path) & PROVIDER_ROOTS)
        if forbidden:
            violations[str(path.relative_to(ROOT))] = forbidden

    assert violations == {}


def test_provider_sdk_is_optional_and_isolated() -> None:
    configuration = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    dependencies = configuration["project"]["dependencies"]
    extras = configuration["project"]["optional-dependencies"]

    assert not any(
        dependency.split("[", 1)[0].startswith("anthropic") for dependency in dependencies
    )
    assert any(dependency.startswith("anthropic") for dependency in extras["legacy-agent"])


def test_prohibited_excel_runtime_and_oracle_modules_are_absent() -> None:
    prohibited = [
        PACKAGE / "excel_oracle.py",
        PACKAGE / "windows_excel_worker.py",
        PACKAGE / "vba_conformance.py",
        PACKAGE / "runtime" / "__init__.py",
        ROOT / "tools" / "windows_excel_oracle.py",
    ]

    assert [str(path.relative_to(ROOT)) for path in prohibited if path.exists()] == []


def test_public_api_does_not_import_legacy_or_choose_a_model() -> None:
    source = (PACKAGE / "api.py").read_text(encoding="utf-8")

    assert "legacy_agent" not in source
    assert "AgentRewriter" not in source
    assert "LLMVBATranslator" not in source
    assert "use_agent: bool = True" not in source
