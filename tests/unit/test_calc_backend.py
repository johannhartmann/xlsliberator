"""Tests for Calc backend discovery and runtime profiles."""

from pathlib import Path
from typing import Any

from xlsliberator import calc_backend
from xlsliberator.calc_backend import create_isolated_user_profile, discover_backends
from xlsliberator.formula_engine import FormulaDialect, FormulaParseResult
from xlsliberator.validation_models import TargetKind


def test_discover_backends_when_missing(monkeypatch: Any) -> None:
    """Missing office executables should return an empty list."""
    monkeypatch.setattr(calc_backend.shutil, "which", lambda _name: None)

    assert discover_backends() == []


def test_discover_backends_parses_libreoffice(monkeypatch: Any) -> None:
    """LibreOffice version strings should produce LibreOffice backend info."""
    monkeypatch.setattr(
        calc_backend.shutil,
        "which",
        lambda name: "/usr/bin/libreoffice" if name == "libreoffice" else None,
    )
    monkeypatch.setattr(calc_backend, "detect_version", lambda _exe: "LibreOffice 7.6.4.1")

    backends = discover_backends()

    assert len(backends) == 1
    assert backends[0].info.kind == TargetKind.LIBREOFFICE
    assert backends[0].info.version == "LibreOffice 7.6.4.1"


def test_discover_backends_parses_openoffice(monkeypatch: Any) -> None:
    """OpenOffice version strings should produce OpenOffice backend info."""
    monkeypatch.setattr(
        calc_backend.shutil,
        "which",
        lambda name: "/usr/bin/openoffice" if name == "openoffice" else None,
    )
    monkeypatch.setattr(calc_backend, "detect_version", lambda _exe: "OpenOffice 4.1.15")

    backends = discover_backends()

    assert len(backends) == 1
    assert backends[0].info.kind == TargetKind.OPENOFFICE
    assert backends[0].info.version == "OpenOffice 4.1.15"


def test_create_isolated_user_profile_url() -> None:
    """Isolated profiles should expose a file URL usable by office."""
    with create_isolated_user_profile("xlsliberator-test-") as profile:
        assert profile.user_installation_dir.exists()
        assert profile.user_installation_url.startswith("file://")
        assert profile.user_installation_arg.startswith("-env:UserInstallation=file://")
        assert isinstance(profile.env, dict)

    assert not Path(profile.user_installation_dir).exists()


def test_backend_formula_parse_hook_uses_uno_when_available(monkeypatch: Any) -> None:
    """Backend formula parse hook should return UNO parser results when available."""
    backend = calc_backend.LibreOfficeBackend("/usr/bin/libreoffice", "LibreOffice 7.6")
    expected = FormulaParseResult(
        success=True,
        formula="=SUM(1;2)",
        dialect=FormulaDialect.CALC_A1,
        tokens=["token"],
        details={"target_parser": "uno_formula_parser"},
    )

    monkeypatch.setattr(calc_backend, "_parse_with_uno_formula_parser", lambda **_kwargs: expected)

    result = backend.parse_formula_text("=SUM(1;2)", sheet_name="Sheet1")

    assert result.success
    assert result.details["target_parser"] == "uno_formula_parser"


def test_backend_formula_parse_hook_reports_structural_fallback(monkeypatch: Any) -> None:
    """Backend formula parse hook should disclose fallback parser scope."""
    backend = calc_backend.LibreOfficeBackend("/usr/bin/libreoffice", "LibreOffice 7.6")

    monkeypatch.setattr(calc_backend, "_parse_with_uno_formula_parser", lambda **_kwargs: None)

    result = backend.parse_formula_text("=SUM(1;2)", sheet_name="Sheet1")

    assert result.success
    assert result.details["backend_kind"] == "libreoffice"
    assert result.details["target_parser"] == "basic_structural_fallback"
    assert "target_parser_unavailable" in result.details


def test_discover_backends_no_phantom_openoffice_when_version_unknown(monkeypatch: Any) -> None:
    """A LibreOffice-only host without a version string must not yield a phantom OpenOffice."""
    monkeypatch.setattr(
        calc_backend.shutil,
        "which",
        lambda name: "/usr/bin/soffice" if name == "soffice" else None,
    )
    monkeypatch.setattr(calc_backend, "detect_version", lambda _exe: None)

    backends = discover_backends()

    assert len(backends) == 1
    assert backends[0].info.kind == TargetKind.LIBREOFFICE
