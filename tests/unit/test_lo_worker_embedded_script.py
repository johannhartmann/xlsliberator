"""Tests for document-local Python execution in the pinned office worker."""

from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from xlsliberator.lo_worker import (
    _execute_embedded_python_script,
    _is_document_python_script,
)


class _Cell:
    String = ""


class _Range:
    def __init__(self, cell: _Cell) -> None:
        self._cell = cell

    def getCellRangeByName(self, _address: str) -> _Cell:
        return self._cell


class _Sheets:
    def __init__(self, cell: _Cell) -> None:
        self._cell = cell

    def getByName(self, _name: str) -> _Range:
        return _Range(self._cell)


class _Document:
    def __init__(self, cell: _Cell) -> None:
        self.Sheets = _Sheets(cell)


def test_document_python_script_executes_with_active_document(tmp_path: Path) -> None:
    workbook = tmp_path / "macro.ods"
    module = (
        "def OnClick(*_args):\n"
        "    document = XSCRIPTCONTEXT.getDocument()\n"
        '    document.Sheets.getByName("Sheet1").getCellRangeByName("D4").String = "fired"\n'
    )
    with ZipFile(workbook, "w") as archive:
        archive.writestr("Scripts/python/CertificationControl.py", module)
    uri = "vnd.sun.star.script:CertificationControl.py$OnClick?language=Python&location=document"
    cell = _Cell()

    result = _execute_embedded_python_script(
        {"script_uri": uri, "ods_path": str(workbook)},
        {"desktop": object(), "component_context": object()},
        _Document(cell),
    )

    assert result["executed"] is True
    assert result["executor"] == "libreoffice-bundled-python"
    assert cell.String == "fired"


def test_only_document_local_python_uri_uses_embedded_executor() -> None:
    assert _is_document_python_script(
        "vnd.sun.star.script:Module.py$run?language=Python&location=document"
    )
    assert not _is_document_python_script(
        "vnd.sun.star.script:Module.py$run?language=Python&location=user"
    )
    assert not _is_document_python_script(
        "vnd.sun.star.script:Standard.Module.run?language=Basic&location=document"
    )
