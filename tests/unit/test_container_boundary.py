"""Tests for the Docker-only application startup boundary."""

from __future__ import annotations

from typing import Any

import pytest
from click.testing import CliRunner

from xlsliberator.api import ConversionError


def test_conversion_rejects_host_before_constructing_office_runtime(monkeypatch: Any) -> None:
    from xlsliberator import api, primitives

    monkeypatch.setattr(
        "xlsliberator.container_boundary.application_container_is_authorized",
        lambda: False,
    )

    def forbidden_runtime(*_args: Any, **_kwargs: Any) -> Any:
        raise AssertionError("office runtime must not be constructed on the host")

    monkeypatch.setattr(primitives, "LibreOfficeDockerRuntime", forbidden_runtime)
    with pytest.raises(ConversionError, match="Host Python execution is forbidden"):
        api.convert_native(api.Path("input.xlsx"), api.Path("output.ods"))


def test_application_container_requires_marker_and_docker(monkeypatch: Any) -> None:
    from xlsliberator import container_boundary

    monkeypatch.setenv("XLSLIBERATOR_APPLICATION_CONTAINER", "1")
    monkeypatch.setattr(container_boundary.Path, "is_file", lambda _path: False)
    assert container_boundary.application_container_is_authorized() is False


def test_cli_command_rejects_unauthorized_container_without_traceback(monkeypatch: Any) -> None:
    from xlsliberator.cli import cli

    monkeypatch.setattr(
        "xlsliberator.container_boundary.application_container_is_authorized",
        lambda: False,
    )
    result = CliRunner().invoke(cli, ["cleanup-jobs", "--data-dir", "/tmp"])
    assert result.exit_code != 0
    assert "Host Python execution is forbidden" in result.output
    assert "Traceback" not in result.output


def test_repository_instructions_forbid_local_python_and_office() -> None:
    from pathlib import Path

    instructions = (Path(__file__).parents[2] / "AGENTS.md").read_text(encoding="utf-8")
    assert "Never start host Python" in instructions
    assert "There is no host executable discovery" in instructions
    assert "uv run pytest" not in instructions
    assert "simple conversion paths may still use `soffice`" not in instructions
