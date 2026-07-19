"""Architecture guards for the XLSLiberator runtime."""

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
        if (PACKAGE / "open_swe_agent") in path.parents:
            continue
        forbidden = sorted(_imports(path) & PROVIDER_ROOTS)
        if forbidden:
            violations[str(path.relative_to(ROOT))] = forbidden

    assert violations == {}


def test_provider_sdks_are_not_package_dependencies() -> None:
    configuration = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    dependencies = configuration["project"]["dependencies"]

    assert not any(
        dependency.split("[", 1)[0].startswith("anthropic") for dependency in dependencies
    )
    assert "legacy-agent" not in configuration["project"]["optional-dependencies"]


def test_open_swe_is_the_only_agent_orchestrator_surface() -> None:
    assert not list((PACKAGE / "legacy_agent").glob("*.py"))
    assert not list((PACKAGE / "orchestrator").glob("*.py"))
    assert not (PACKAGE / "web" / "orchestrator.py").exists()
    assert (PACKAGE / "web" / "open_swe.py").is_file()
    assert (PACKAGE / "open_swe_agent" / "graph.py").is_file()

    compose = (ROOT / "docker-compose.yml").read_text(encoding="utf-8")
    runner = (PACKAGE / "web" / "runner.py").read_text(encoding="utf-8")
    assert "xlsliberator-orchestrator" not in compose
    assert "test-orchestrator" not in compose
    assert "ci-orchestrator" not in compose
    assert "XLSLIBERATOR_ORCHESTRATOR_" not in compose
    assert "XLSLIBERATOR_OPEN_SWE_URL" in compose
    assert "xlsliberator-open-swe:" in compose
    assert "docker/open-swe/Dockerfile" in compose
    assert "OpenSWEClient" in runner
    assert "local conversion is disabled" in runner


def test_open_swe_has_no_shell_backend_or_automatic_paid_model() -> None:
    graph = (PACKAGE / "open_swe_agent" / "graph.py").read_text(encoding="utf-8")

    assert "FilesystemBackend" in graph
    assert "StateBackend" in graph
    assert "LocalShellBackend" not in graph
    assert "DockerBackend" not in graph
    assert "ModelFallbackMiddleware" not in graph
    assert "use_gateway=False" in graph
    assert "XLSLIBERATOR_OPEN_SWE_MODEL is required" in graph
    assert "XLSLIBERATOR_GITHUB_MODELS_ENABLED" in graph
    assert "github_models:" in graph


def test_open_swe_image_uses_verified_upstream_source_and_lockfile() -> None:
    dockerfile = (ROOT / "docker" / "open-swe" / "Dockerfile").read_text(encoding="utf-8")
    compose = (ROOT / "docker-compose.yml").read_text(encoding="utf-8")

    assert "OPEN_SWE_COMMIT" in dockerfile
    assert "OPEN_SWE_ARCHIVE_SHA256" in dockerfile
    assert "sha256sum -c -" in dockerfile
    assert "uv sync --frozen --no-dev" in dockerfile
    assert "LANGGRAPH_CLI_NO_ANALYTICS=1" in dockerfile
    assert "LANGGRAPH_NO_VERSION_CHECK=1" in dockerfile
    assert (
        "/var/run/docker.sock"
        not in compose.split("  xlsliberator-open-swe:\n", 1)[1].split("\n  xlsliberator-web:", 1)[
            0
        ]
    )


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
    assert "use_agent" not in source
