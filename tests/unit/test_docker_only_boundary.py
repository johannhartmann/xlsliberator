"""Regression tests for the absolute host-office prohibition."""

from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import Any

import pytest

from xlsliberator.uno_conn import UnoConnectionError, UnoCtx


def test_legacy_uno_context_fails_without_importing_or_starting_office(monkeypatch: Any) -> None:
    calls: list[str] = []

    def forbidden_import(name: str, *_args: Any, **_kwargs: Any) -> Any:
        calls.append(name)
        raise AssertionError(name)

    monkeypatch.setattr("builtins.__import__", forbidden_import)
    with pytest.raises(UnoConnectionError, match="Docker runtime"):
        UnoCtx().connect()
    assert calls == []


def test_pyuno_imports_exist_only_in_container_worker() -> None:
    source_root = Path(__file__).parents[2] / "src" / "xlsliberator"
    violations: list[str] = []
    for path in source_root.glob("*.py"):
        if path.name == "lo_worker.py":
            continue
        tree = ast.parse(path.read_text())
        for node in ast.walk(tree):
            if isinstance(node, ast.Import) and any(alias.name == "uno" for alias in node.names):
                violations.append(f"{path.name}:{node.lineno}")
            if isinstance(node, ast.ImportFrom) and node.module == "uno":
                violations.append(f"{path.name}:{node.lineno}")
    assert violations == []


def test_container_worker_has_fail_closed_host_guard() -> None:
    """Prove the guard statically without importing the container worker on host."""
    source_root = Path(__file__).parents[2] / "src" / "xlsliberator"
    source = (source_root / "lo_worker.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    functions = {node.name: node for node in tree.body if isinstance(node, ast.FunctionDef)}
    assert "_require_office_container" in functions
    for function_name in ("main", "_dispatch"):
        calls = [
            node
            for node in ast.walk(functions[function_name])
            if isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "_require_office_container"
        ]
        assert calls, f"{function_name} must enforce the container guard"
    assert "XLSLIBERATOR_OFFICE_CONTAINER" in source
    assert 'Path("/.dockerenv").is_file()' in source
    assert "OFFICE_PYTHON_PREFIX" in source
    assert "forbidden on the host" in source


def test_calc_backend_has_no_host_office_execution_path() -> None:
    source_root = Path(__file__).parents[2] / "src" / "xlsliberator"
    source = (source_root / "calc_backend.py").read_text(encoding="utf-8")
    tree = ast.parse(source)
    imported_modules = {
        alias.name
        for node in ast.walk(tree)
        if isinstance(node, ast.Import)
        for alias in node.names
    }
    imported_from = {node.module for node in ast.walk(tree) if isinstance(node, ast.ImportFrom)}
    assert "subprocess" not in imported_modules
    assert "xlsliberator.lo_worker_client" not in imported_from
    assert "soffice" not in source.lower()


def test_python_startup_guard_blocks_uno_outside_office_container() -> None:
    """The unit-test container is Docker, but it is not the office runtime."""
    import subprocess
    import sys

    result = subprocess.run(
        [sys.executable, "-c", "import uno"],
        cwd=Path(__file__).parents[2],
        capture_output=True,
        text=True,
        check=False,
    )
    assert result.returncode != 0
    assert "forbidden outside the pinned LibreOffice Docker runtime" in result.stderr


def test_project_commands_do_not_execute_python_on_the_host() -> None:
    root = Path(__file__).parents[2]
    makefile = (root / "Makefile").read_text(encoding="utf-8")
    compose = (root / "docker-compose.yml").read_text(encoding="utf-8")
    readme = (root / "README.md").read_text(encoding="utf-8")
    examples = (root / "examples/README.md").read_text(encoding="utf-8")
    assert "DOCKER_TEST := docker compose run --rm test" in makefile
    assert "dockerfile: docker/test/Dockerfile" in compose
    assert "sudo apt-get install python" not in readme
    assert re.search(r"(?m)^\s*(?:npx|npm)\s", examples) is None
    assert "\ncurl " not in examples
    for command in ("pytest", "ruff", "mypy", "pip-audit", "bandit"):
        assert f"\t{command} " not in makefile


def test_office_source_wrapper_runs_python_only_through_docker() -> None:
    root = Path(__file__).parents[2]
    wrapper = (root / "tools" / "office").read_text(encoding="utf-8")
    logical_lines = wrapper.replace("\\\n", " ").splitlines()
    python_commands = [line.strip() for line in logical_lines if "python3 tools/office.py" in line]

    assert len(python_commands) == 2
    assert all(command.startswith("docker compose ") for command in python_commands)


def test_office_source_fetch_uses_release_tarball_make_target() -> None:
    root = Path(__file__).parents[2]
    office_tool = (root / "tools" / "office.py").read_text(encoding="utf-8")

    assert '_run(["make", "fetch"], cwd=download_tree)' in office_tool
    assert '_run(["./download"]' not in office_tool


def test_office_source_build_propagates_requested_parallelism() -> None:
    root = Path(__file__).parents[2]
    office_tool = (root / "tools" / "office.py").read_text(encoding="utf-8")

    assert 'f"--with-parallelism={parallelism}"' in office_tool
    assert "_configure(source, manifest, parallelism=args.jobs)" in office_tool


def test_office_source_build_uses_pinned_bundled_nss_crypto_backend() -> None:
    root = Path(__file__).parents[2]
    manifest = (root / "office/libreoffice/manifest.json").read_text(encoding="utf-8")

    assert '"--without-system-nss"' in manifest


def test_runner_uses_shared_workspace_for_pytest_and_office_jobs() -> None:
    root = Path(__file__).parents[2]
    compose = (root / "docker-compose.yml").read_text(encoding="utf-8")
    assert "XLSLIBERATOR_WORKSPACE_ROOTS: ${PWD}:/tmp/pytest-tmp" in compose
    assert "PYTEST_ADDOPTS: --basetemp=/tmp/pytest-tmp" in compose
    assert "XLSLIBERATOR_RUNTIME_TEMP_ROOT: ${PWD}/artifacts/runtime-tmp" in compose
    assert "XLSLIBERATOR_OPEN_SWE_WORKSPACE_ROOT: ${PWD}/artifacts/open-swe-workspaces" in compose


def test_only_mcp_is_a_trusted_docker_execution_gateway() -> None:
    root = Path(__file__).parents[2]
    compose = (root / "docker-compose.yml").read_text(encoding="utf-8")
    dockerfile = (root / "Dockerfile").read_text(encoding="utf-8")
    open_swe_service = compose.split("  xlsliberator-open-swe:\n", 1)[1].split(
        "\n  xlsliberator-web:", 1
    )[0]
    web_service = compose.split("  xlsliberator-web:\n", 1)[1].split("\n  xlsliberator-mcp:", 1)[0]
    mcp_service = compose.split("  xlsliberator-mcp:\n", 1)[1].split("\nvolumes:", 1)[0]
    runtime_service = compose.split("  libreoffice-runtime:\n", 1)[1].split(
        "\n  office-source-fetch:", 1
    )[0]

    assert "docker-cli" in dockerfile
    assert "/var/run/docker.sock" not in open_swe_service
    assert "/var/run/docker.sock" not in web_service
    assert "/var/run/docker.sock:/var/run/docker.sock" in mcp_service
    assert "XLSLIBERATOR_OPEN_SWE_MODEL" in open_swe_service
    assert "XLSLIBERATOR_GITHUB_MODELS_ENABLED" in open_swe_service
    assert "XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT" not in web_service
    assert "XLSLIBERATOR_OFFICE_CONTAINER" not in web_service
    assert "XLSLIBERATOR_OPEN_SWE_URL" in web_service
    assert "XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT" in mcp_service
    assert "/var/run/docker.sock" not in runtime_service


def test_office_image_registers_its_non_root_runtime_identity() -> None:
    root = Path(__file__).parents[2]
    dockerfile = (root / "docker/office/libreoffice/Dockerfile").read_text(encoding="utf-8")

    assert "groupadd --gid 10001 xlsliberator" in dockerfile
    assert "--uid 10001" in dockerfile
    assert "--gid 10001" in dockerfile
    assert "--home-dir /tmp/home" in dockerfile
    assert "--shell /usr/sbin/nologin" in dockerfile
    assert "USER 10001:10001" in dockerfile


def test_application_image_registers_the_shared_non_root_service_identity() -> None:
    root = Path(__file__).parents[2]
    dockerfile = (root / "Dockerfile").read_text(encoding="utf-8")

    assert "groupadd --gid 10001 appuser" in dockerfile
    assert "--uid 10001" in dockerfile
    assert "--gid 10001" in dockerfile
    assert "--shell /usr/sbin/nologin" in dockerfile
    assert "USER 10001:10001" in dockerfile


def test_gui_entrypoint_starts_openbox_on_the_private_display() -> None:
    root = Path(__file__).parents[2]
    entrypoint = (root / "docker/office/gui/runtime-entrypoint").read_text(encoding="utf-8")

    assert 'export DISPLAY="$display"' in entrypoint
    assert 'openbox >"$XLSLIBERATOR_OPENBOX_LOG" 2>&1 &' in entrypoint
    assert "openbox --display" not in entrypoint


def test_gui_image_uses_generic_x11_without_gtk_or_acceleration() -> None:
    root = Path(__file__).parents[2]
    dockerfile = (root / "docker/office/gui/Dockerfile").read_text(encoding="utf-8")

    assert "libgtk-3-0" not in dockerfile
    assert "GDK_BACKEND" not in dockerfile
    assert "NO_AT_BRIDGE" not in dockerfile
    assert "SAL_DISABLE_CUPS=true" in dockerfile
    assert "SAL_DISABLEGL=1" in dockerfile
    assert "SAL_DISABLE_OPENCL=1" in dockerfile
    assert "SAL_DISABLESKIA=1" in dockerfile
    assert "SAL_NO_MOUSEGRABS=1" in dockerfile
    assert "SAL_SYNCHRONIZE=1" in dockerfile
    assert "SAL_USE_VCLPLUGIN=gen" in dockerfile


def test_dependency_audit_has_network_without_office_or_docker_socket_access() -> None:
    root = Path(__file__).parents[2]
    compose = (root / "docker-compose.yml").read_text(encoding="utf-8")
    makefile = (root / "Makefile").read_text(encoding="utf-8")
    scheduled_workflow = (root / ".github/workflows/security.yml").read_text(encoding="utf-8")
    audit_service = compose.split("  security-audit:\n", 1)[1].split("\n  libreoffice-runtime:", 1)[
        0
    ]

    assert "network_mode: bridge" in audit_service
    assert "docker.sock" not in audit_service
    assert "XLSLIBERATOR_OFFICE_CONTAINER" not in audit_service
    assert "$(DOCKER_SECURITY) pip-audit --desc" in makefile
    assert "run --rm security-audit" in scheduled_workflow
    assert "run --rm test pip-audit" not in scheduled_workflow


def test_package_gate_is_offline_and_build_backend_is_in_test_image() -> None:
    root = Path(__file__).parents[2]
    app_dockerfile = (root / "Dockerfile").read_text(encoding="utf-8")
    test_dockerfile = (root / "docker/test/Dockerfile").read_text(encoding="utf-8")
    ci_check = (root / "tools/ci_check.py").read_text(encoding="utf-8")

    assert "build hatchling pip-audit" in test_dockerfile
    assert 'pip install --no-cache-dir ".[web]"' in app_dockerfile
    assert 'pip install --no-cache-dir ".[web,dev]"' in test_dockerfile
    assert "pip install --no-cache-dir -e" not in app_dockerfile
    assert "pip install --no-cache-dir -e" not in test_dockerfile
    assert "cmp -s /app/src/sitecustomize.py" in app_dockerfile
    assert "cmp -s /build/src/sitecustomize.py" in test_dockerfile
    assert "site.getsitepackages()[0]" in app_dockerfile
    assert "site.getsitepackages()[0]" in test_dockerfile
    assert "')/sitecustomize.py\"" in app_dockerfile
    assert "')/sitecustomize.py\"" in test_dockerfile
    assert '"build", "--no-isolation"' in ci_check
    assert '("*.whl", "*.tar.gz")' in ci_check
    assert 'archive.read("sitecustomize.py")' in ci_check
    assert "packaged_guard != expected_guard" in ci_check
