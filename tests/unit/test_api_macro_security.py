"""Tests for safe macro security defaults in conversion."""

import zipfile
from pathlib import Path
from typing import Any

from xlsliberator import api
from xlsliberator.ir_models import ExtractionStats, WorkbookIR


class _DummyCtx:
    def __enter__(self) -> "_DummyCtx":
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        return None


def _patch_minimal_convert(monkeypatch: Any, output_path: Path) -> None:
    def write_ods(_input: Path, _output: Path) -> None:
        with zipfile.ZipFile(output_path, "w") as archive:
            archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")

    monkeypatch.setattr(api, "convert_native", write_ods)
    monkeypatch.setattr(
        api,
        "extract_workbook",
        lambda path: (WorkbookIR(file_path=str(path), file_format="xlsx"), ExtractionStats()),
    )


def test_convert_default_does_not_set_global_macro_security(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Default conversion must not mutate global macro security."""
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    input_path.write_text("placeholder")
    calls: list[int] = []

    _patch_minimal_convert(monkeypatch, output_path)
    import xlsliberator.uno_conn as uno_conn

    monkeypatch.setattr(uno_conn, "UnoCtx", _DummyCtx)
    monkeypatch.setattr(
        uno_conn, "set_macro_security_level", lambda _ctx, level: calls.append(level)
    )

    report = api.convert(input_path, output_path, embed_macros=False)

    assert report.success
    assert calls == []


def test_convert_explicit_opt_in_cannot_mutate_global_macro_security(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    """Even explicit opt-in cannot escape the Docker-only runtime boundary."""
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    input_path.write_text("placeholder")
    calls: list[int] = []

    _patch_minimal_convert(monkeypatch, output_path)
    import xlsliberator.uno_conn as uno_conn

    monkeypatch.setattr(uno_conn, "UnoCtx", _DummyCtx)
    monkeypatch.setattr(
        uno_conn, "set_macro_security_level", lambda _ctx, level: calls.append(level)
    )

    report = api.convert(
        input_path,
        output_path,
        embed_macros=False,
        allow_global_macro_security_change=True,
    )

    assert report.success
    assert calls == []
    assert any("unsupported in the Docker-only runtime" in item for item in report.warnings)
