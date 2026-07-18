"""Tests for provider-neutral deterministic migration primitives."""

from __future__ import annotations

import zipfile
from pathlib import Path

from xlsliberator.ir_models import WorkbookIR
from xlsliberator.primitives import (
    inspect_source_workbook,
    native_convert_workbook,
    validate_ods_package,
)
from xlsliberator.validation_models import GateExecutionStatus, WorkbookArtifactIR


def _write_minimal_ods(path: Path) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            zipfile.ZipInfo("mimetype"),
            b"application/vnd.oasis.opendocument.spreadsheet",
            compress_type=zipfile.ZIP_STORED,
        )
        archive.writestr(
            "META-INF/manifest.xml",
            b'<?xml version="1.0"?><manifest:manifest '
            b'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>',
        )


def test_validate_ods_package_requires_structural_invariants(tmp_path: Path) -> None:
    valid = tmp_path / "valid.ods"
    _write_minimal_ods(valid)

    result = validate_ods_package(valid)

    assert result.status is GateExecutionStatus.PASSED
    assert result.success is True
    assert result.mimetype_first is True
    assert result.mimetype_stored is True


def test_validate_ods_package_fails_missing_manifest(tmp_path: Path) -> None:
    invalid = tmp_path / "invalid.ods"
    with zipfile.ZipFile(invalid, "w") as archive:
        archive.writestr(
            zipfile.ZipInfo("mimetype"),
            b"application/vnd.oasis.opendocument.spreadsheet",
            compress_type=zipfile.ZIP_STORED,
        )

    result = validate_ods_package(invalid)

    assert result.status is GateExecutionStatus.FAILED
    assert "ODS manifest is missing" in result.errors


def test_source_inspection_wraps_inventory_in_typed_result(
    monkeypatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "book.xlsx"
    source.write_bytes(b"source")
    inventory = WorkbookArtifactIR(workbook=WorkbookIR(file_path=str(source), file_format="xlsx"))
    monkeypatch.setattr(
        "xlsliberator.primitives.inspect_workbook",
        lambda _path, role: inventory,
    )

    result = inspect_source_workbook(source)

    assert result.status is GateExecutionStatus.PASSED
    assert result.inventory is inventory


def test_native_convert_reports_runtime_unavailable(monkeypatch, tmp_path: Path) -> None:
    from xlsliberator.docker_runtime import DockerRuntimeUnavailable

    source = tmp_path / "book.xlsx"
    destination = tmp_path / "book.ods"
    source.write_bytes(b"source")

    class UnavailableRuntime:
        def __init__(self, *, timeout_seconds: int) -> None:
            assert timeout_seconds == 120

        def convert(self, _source: Path, _destination: Path):
            raise DockerRuntimeUnavailable("pinned office image unavailable")

    monkeypatch.setattr("xlsliberator.primitives.require_application_container", lambda: None)
    monkeypatch.setattr("xlsliberator.primitives.LibreOfficeDockerRuntime", UnavailableRuntime)

    result = native_convert_workbook(source, destination)

    assert result.status is GateExecutionStatus.UNAVAILABLE
    assert result.success is False
    assert result.errors == ["pinned office image unavailable"]
