"""LibreOffice Python worker for UNO operations.

This module is intentionally standard-library only so it can run under
LibreOffice's bundled Python wrapper.
"""

from __future__ import annotations

import hashlib
import importlib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tempfile
import time
import traceback
import uuid
from collections.abc import Callable
from contextlib import suppress
from datetime import UTC, date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, unquote
from zipfile import ZipFile

DEFAULT_START_TIMEOUT_SECONDS = 20
OFFICE_CONTAINER_MARKER = "XLSLIBERATOR_OFFICE_CONTAINER"
OFFICE_PYTHON_PREFIX = "/opt/libreoffice26.2/program/"
SOURCE_RUNTIME_PREFIX = "/opt/libreoffice/program"


def _authorized_office_prefix() -> str:
    candidate = os.environ.get("XLSLIBERATOR_OFFICE_PYTHON_PREFIX", OFFICE_PYTHON_PREFIX)
    if candidate == OFFICE_PYTHON_PREFIX:
        return candidate
    resolved = str(Path(candidate).resolve())
    if (
        os.environ.get("XLSLIBERATOR_SOURCE_BUILD_CONTAINER") == "1"
        and resolved.startswith("/office-work/worktrees/")
        and resolved.endswith("/instdir/program")
    ):
        return f"{resolved}/"
    if (
        os.environ.get("XLSLIBERATOR_SOURCE_RUNTIME_CONTAINER") == "1"
        and resolved == SOURCE_RUNTIME_PREFIX
    ):
        return f"{resolved}/"
    return OFFICE_PYTHON_PREFIX


def _authorized_office_executable() -> str:
    executable = str(
        Path(
            os.environ.get("XLSLIBERATOR_OFFICE_EXECUTABLE", f"{OFFICE_PYTHON_PREFIX}soffice")
        ).resolve()
    )
    if not executable.startswith(_authorized_office_prefix()):
        raise RuntimeError("Office executable is outside the authorized Docker runtime prefix")
    return executable


def _require_office_container() -> None:
    """Refuse every worker operation outside the pinned office container."""
    python_executable = str(Path(sys.executable).resolve())
    if (
        os.environ.get(OFFICE_CONTAINER_MARKER) != "1"
        or not Path("/.dockerenv").is_file()
        or not python_executable.startswith(_authorized_office_prefix())
    ):
        raise RuntimeError(
            "LibreOffice worker execution is forbidden on the host; use LibreOfficeDockerRuntime"
        )


def main() -> int:
    """Read one JSON request from stdin and write one JSON response to stdout."""
    try:
        _require_office_container()
    except Exception as exc:
        _write_response(
            {
                "success": False,
                "op": "host_execution_forbidden",
                "data": {},
                "error": _error_payload(exc),
            }
        )
        return 78
    try:
        request = json.loads(sys.stdin.read() or "{}")
    except Exception as exc:
        _write_response(
            {
                "success": False,
                "op": "unknown",
                "data": {},
                "error": _error_payload(exc),
            }
        )
        return 1

    op = str(request.get("op", "unknown"))
    try:
        data = _dispatch(request)
        _write_response({"success": True, "op": op, "data": data, "error": None})
        return 0
    except Exception as exc:
        _write_response(
            {
                "success": False,
                "op": op,
                "data": {},
                "error": _error_payload(exc),
            }
        )
        return 1


def _dispatch(request: dict[str, Any]) -> dict[str, Any]:
    _require_office_container()
    op = str(request.get("op", ""))
    if op == "ping":
        import uno

        return {
            "uno_importable": True,
            "uno_module": getattr(uno, "__file__", None),
            "python_executable": sys.executable,
        }
    if op == "runtime_probe":
        return _runtime_probe()
    if op == "convert_document":
        return _convert_document(request)
    if op == "create_controls_fixture":
        return _create_controls_fixture(request)
    if op == "validate_python":
        return _validate_python(request)
    if op == "validate_document":
        return _validate_document(request)
    if op == "run_scenario":
        return _run_scenario(request)
    if op == "inspect_document_cells":
        return _with_document(request, _inspect_document_cells)
    if op == "list_formula_cells":
        return _with_document(request, _list_formula_cells)
    if op == "apply_document_repairs":
        return _apply_document_repairs(request)
    if op == "parse_formula":
        return _parse_formula(request)
    if op == "evaluate_formula":
        return _evaluate_formula(request)
    if op == "list_sheets":
        return _with_document(request, _list_sheets)
    if op == "read_cell":
        return _with_document(request, _read_cell)
    if op == "get_sheet_data":
        return _with_document(request, _get_sheet_data)
    if op == "get_cell_colors":
        return _with_document(request, _get_cell_colors)
    if op == "execute_script":
        return _with_document(request, _execute_script)
    if op == "execute_button_handler":
        return _with_document(request, _execute_button_handler)
    if op == "recalculate_document":
        return _with_document(request, _recalculate_document)
    raise ValueError(f"Unsupported worker op: {op}")


def _runtime_probe() -> dict[str, Any]:
    """Return the office/Python/PyUNO identity from inside the runtime image."""
    import uno

    pyuno = importlib.import_module("pyuno")

    executable = _authorized_office_executable()
    result = subprocess.run(
        [executable, "--version"], capture_output=True, text=True, timeout=10, check=True
    )
    version_output = (result.stdout or result.stderr).strip()
    build = version_output.split()[1] if len(version_output.split()) > 1 else ""
    uno_path = Path(str(uno.__file__)).resolve()
    pyuno_path = Path(str(pyuno.__file__)).resolve()
    wrapper_path = Path(
        os.environ.get("XLSLIBERATOR_WORKER_WRAPPER", "/usr/local/bin/runtime-entrypoint")
    )
    return {
        "libreoffice_build": build,
        "runtime_variant": os.environ.get("XLSLIBERATOR_RUNTIME_VARIANT", ""),
        "source_commit": os.environ.get("XLSLIBERATOR_SOURCE_COMMIT", ""),
        "patch_set_sha256": os.environ.get("XLSLIBERATOR_PATCH_SET_SHA256", ""),
        "office_program_prefix": str(_authorized_office_prefix()),
        "version_output": version_output,
        "office_executable": executable,
        "office_sha256": _sha256_file(Path(executable)),
        "python_executable": sys.executable,
        "python_version": platform.python_version(),
        "uno_importable": True,
        "uno_module": str(uno_path),
        "uno_module_sha256": _sha256_file(uno_path),
        "pyuno_native_module": str(pyuno_path),
        "pyuno_native_sha256": _sha256_file(pyuno_path),
        "pythonpath": os.environ.get("PYTHONPATH", ""),
        "worker_wrapper": str(wrapper_path),
        "worker_wrapper_sha256": _sha256_file(wrapper_path),
        "base_image_digest": os.environ.get("XLSLIBERATOR_BASE_IMAGE_DIGEST", ""),
        "installed_package_manifest": _installed_package_manifest(),
        "architecture": platform.machine(),
    }


def _installed_package_manifest() -> list[dict[str, str]]:
    """Return a deterministic manifest of every installed Debian package."""
    result = subprocess.run(
        ["dpkg-query", "-W", "-f=${Package}\t${Version}\t${Architecture}\n"],
        capture_output=True,
        text=True,
        timeout=10,
        check=True,
    )
    manifest = []
    for line in sorted(result.stdout.splitlines()):
        name, version, architecture = line.split("\t", 2)
        manifest.append({"name": name, "version": version, "architecture": architecture})
    return manifest


def _convert_document(request: dict[str, Any]) -> dict[str, Any]:
    """Convert one staged input workbook to ODS inside a fresh office session."""
    input_path = Path(str(request["input_path"])).resolve()
    output_path = Path(str(request["output_path"])).resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            input_url = session["uno"].systemPathToFileUrl(str(input_path))
            output_url = session["uno"].systemPathToFileUrl(str(output_path))
            document = session["desktop"].loadComponentFromURL(
                input_url, "_blank", 0, (_property_value("Hidden", True),)
            )
            if document is None:
                raise RuntimeError(f"LibreOffice could not open input: {input_path}")
            document.storeToURL(
                output_url,
                (
                    _property_value("FilterName", "calc8"),
                    _property_value("Overwrite", True),
                ),
            )
        finally:
            _close_document(document, save=False)
    if not output_path.is_file():
        raise RuntimeError("LibreOffice returned without producing the ODS output")
    return {
        "output_path": str(output_path),
        "output_sha256": _sha256_file(output_path),
        "output_size": output_path.stat().st_size,
    }


def _create_controls_fixture(request: dict[str, Any]) -> dict[str, Any]:
    """Create a real ODS form control with a document-local event handler."""
    output_path = Path(str(request["output_path"])).resolve()
    button_name = str(request.get("button_name") or "CertificationButton")
    marker_address = str(request.get("marker_address") or "D4")
    marker_value = str(request.get("marker_value") or "control-event-fired")
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            document = session["desktop"].loadComponentFromURL(
                "private:factory/scalc",
                "_blank",
                0,
                (_property_value("Hidden", True),),
            )
            if document is None:
                raise RuntimeError("LibreOffice could not create the controls fixture")
            sheet = document.getSheets().getByIndex(0)
            sheet.setName("Sheet1")
            sheet.getCellRangeByName("A1").setString("Controls certification fixture")
            document.storeAsURL(
                session["uno"].systemPathToFileUrl(str(output_path)),
                (
                    _property_value("FilterName", "calc8"),
                    _property_value("Overwrite", True),
                ),
            )

            draw_page = sheet.getDrawPage()
            forms = draw_page.getForms()
            form = document.createInstance("com.sun.star.form.component.DataForm")
            form.Name = "CertificationForm"
            forms.insertByIndex(0, form)

            shape = document.createInstance("com.sun.star.drawing.ControlShape")
            position = session["uno"].createUnoStruct("com.sun.star.awt.Point")
            position.X = 1000
            position.Y = 1000
            size = session["uno"].createUnoStruct("com.sun.star.awt.Size")
            size.Width = 5000
            size.Height = 1200
            shape.setPosition(position)
            shape.setSize(size)

            button = document.createInstance("com.sun.star.form.component.CommandButton")
            button.Name = button_name
            button.Label = "Run certification event"
            shape.setControl(button)
            form.insertByIndex(0, button)
            draw_page.add(shape)

            document.store()
        finally:
            _close_document(document, save=False)
    if not output_path.is_file():
        raise RuntimeError("LibreOffice did not produce the controls fixture")
    return {
        "output_path": str(output_path),
        "output_sha256": _sha256_file(output_path),
        "button_name": button_name,
        "marker_address": marker_address,
        "marker_value": marker_value,
    }


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _validate_python(request: dict[str, Any]) -> dict[str, Any]:
    """Compile translated code with the runtime that provides matching PyUNO."""
    import uno
    import unohelper

    code = str(request.get("python_code", ""))
    try:
        compile(code, "<translated-macro>", "exec")
    except SyntaxError as exc:
        return {
            "compatible": False,
            "error": f"{exc.msg} at line {exc.lineno}",
            "uno_module": getattr(uno, "__file__", None),
            "unohelper_module": getattr(unohelper, "__file__", None),
        }
    return {
        "compatible": True,
        "uno_module": getattr(uno, "__file__", None),
        "unohelper_module": getattr(unohelper, "__file__", None),
        "python_executable": sys.executable,
    }


def _validate_document(request: dict[str, Any]) -> dict[str, Any]:
    """Exercise open/recalculate/save/close/reopen without mutating the input."""
    source = Path(str(request["ods_path"])).resolve()
    before_sha = _sha256_file(source)
    job_root = Path(os.environ.get("XLSLIBERATOR_JOB_DIR", "/job")).resolve()
    job_root.mkdir(parents=True, exist_ok=True)
    saved = job_root / f"validated-{uuid.uuid4().hex}.ods"
    stages: dict[str, dict[str, Any]] = {
        name: {"status": "not_run", "error": None}
        for name in ("open", "recalculate", "save", "close", "reopen", "package")
    }
    runtime: dict[str, Any] = {}
    with _office_session(request, use_gui=False) as session:
        runtime = {
            "profile_identifier": session["profile_identifier"],
            "profile_path": session["profile_path"],
            "office_executable": session["office_executable"],
            "office_pid": session["office_pid"],
        }
        document = None
        try:
            source_url = session["uno"].systemPathToFileUrl(str(source))
            document = session["desktop"].loadComponentFromURL(
                source_url, "_blank", 0, (_property_value("Hidden", True),)
            )
            if document is None:
                raise RuntimeError("LibreOffice returned no document")
            stages["open"]["status"] = "passed"
        except Exception as exc:
            stages["open"] = {"status": "failed", "error": str(exc)}

        if document is not None:
            try:
                document.calculateAll()
                stages["recalculate"]["status"] = "passed"
            except Exception as exc:
                stages["recalculate"] = {"status": "failed", "error": str(exc)}
            try:
                saved_url = session["uno"].systemPathToFileUrl(str(saved))
                document.storeAsURL(
                    saved_url,
                    (
                        _property_value("FilterName", "calc8"),
                        _property_value("Overwrite", True),
                    ),
                )
                stages["save"]["status"] = "passed"
            except Exception as exc:
                stages["save"] = {"status": "failed", "error": str(exc)}
            try:
                document.close(True)
                document = None
                stages["close"]["status"] = "passed"
            except Exception as exc:
                stages["close"] = {"status": "failed", "error": str(exc)}

        reopened = None
        if stages["save"]["status"] == "passed":
            try:
                saved_url = session["uno"].systemPathToFileUrl(str(saved))
                reopened = session["desktop"].loadComponentFromURL(
                    saved_url, "_blank", 0, (_property_value("Hidden", True),)
                )
                if reopened is None:
                    raise RuntimeError("LibreOffice returned no reopened document")
                stages["reopen"]["status"] = "passed"
            except Exception as exc:
                stages["reopen"] = {"status": "failed", "error": str(exc)}
            finally:
                _close_document(reopened, save=False)
        _close_document(document, save=False)

    try:
        if not saved.is_file() or not _is_valid_ods_package(saved):
            raise RuntimeError("Saved output is not a structurally valid ODS package")
        stages["package"]["status"] = "passed"
    except Exception as exc:
        stages["package"] = {"status": "failed", "error": str(exc)}

    after_sha = _sha256_file(source)
    return {
        "stages": stages,
        "runtime": runtime,
        "source_sha256_before": before_sha,
        "source_sha256_after": after_sha,
        "source_mutated": before_sha != after_sha,
        "saved_sha256": _sha256_file(saved) if saved.is_file() else None,
    }


def _is_valid_ods_package(path: Path) -> bool:
    try:
        with ZipFile(path) as archive:
            names = set(archive.namelist())
            return "mimetype" in names and "content.xml" in names
    except (OSError, ValueError):
        return False


def _parse_formula(request: dict[str, Any]) -> dict[str, Any]:
    formula = str(request["formula"])
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            desktop = session["desktop"]
            hidden = _property_value("Hidden", True)
            ods_path = request.get("ods_path")
            document_url = (
                session["uno"].systemPathToFileUrl(str(Path(str(ods_path)).resolve()))
                if ods_path
                else "private:factory/scalc"
            )
            document = desktop.loadComponentFromURL(document_url, "_blank", 0, (hidden,))
            if document is None:
                raise RuntimeError("LibreOffice did not create a Calc document")

            parser = document.createInstance("com.sun.star.sheet.FormulaParser")
            _set_parser_property(parser, "CompileEnglish", True)
            _set_parser_property(parser, "ParameterSeparator", ";")
            address = _uno_struct("com.sun.star.table.CellAddress")
            sheet_name = request.get("sheet_name")
            cell_address = str(request.get("cell_address") or "A1")
            if sheet_name is not None:
                cell = _get_sheet(document, sheet_name).getCellRangeByName(cell_address)
                address = cell.getCellAddress()

            tokens = parser.parseFormula(formula, address)
            printed = parser.printFormula(tokens, address)
            roundtrip_tokens = parser.parseFormula(printed, address)
            token_strings = [_formula_token_to_string(token) for token in tokens]
            roundtrip_token_strings = [
                _formula_token_to_string(token) for token in roundtrip_tokens
            ]
            syntax_errors = _formula_lexical_errors(formula)
            return {
                "formula": formula,
                "tokens": token_strings,
                "roundtrip_formula": printed,
                "roundtrip_tokens": roundtrip_token_strings,
                "roundtrip_equivalent": token_strings == roundtrip_token_strings,
                "parser_accepted": not syntax_errors,
                "syntax_errors": syntax_errors,
                "sheet_name": sheet_name,
                "cell_address": cell_address,
                "document_context": "target" if ods_path else "blank",
            }
        finally:
            _close_document(document, save=False)


def _evaluate_formula(request: dict[str, Any]) -> dict[str, Any]:
    """Evaluate one minimized formula in a new disposable Calc document."""
    formula = str(request["formula"])
    if not formula.startswith("="):
        raise ValueError("formula must begin with '='")
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            document = session["desktop"].loadComponentFromURL(
                "private:factory/scalc",
                "_blank",
                0,
                (_property_value("Hidden", True),),
            )
            if document is None:
                raise RuntimeError("LibreOffice did not create a Calc document")
            cell = document.getSheets().getByIndex(0).getCellRangeByName("A1")
            cell.setFormula(formula)
            document.calculateAll()
            return {
                "formula": formula,
                "formula_after": str(cell.getFormula()),
                "error_code": int(cell.getError()),
                "string": str(cell.getString()),
                "numeric_value": float(cell.getValue()),
                "cell_type": _cell_type_name(cell.getType()),
            }
        finally:
            _close_document(document, save=False)


DocumentHandler = Callable[[dict[str, Any], dict[str, Any], Any], dict[str, Any]]


def _with_document(request: dict[str, Any], handler: DocumentHandler) -> dict[str, Any]:
    ods_path = Path(str(request["ods_path"])).resolve()
    if not ods_path.exists():
        raise FileNotFoundError(f"File not found: {ods_path}")

    with _office_session(request, use_gui=bool(request.get("use_gui", False))) as session:
        document = None
        try:
            file_url = session["uno"].systemPathToFileUrl(str(ods_path))
            load_props = (
                _property_value("MacroExecutionMode", 4),
                _property_value("Hidden", not bool(request.get("use_gui", False))),
            )
            document = session["desktop"].loadComponentFromURL(file_url, "_blank", 0, load_props)
            if document is None:
                raise RuntimeError(f"LibreOffice could not open document: {ods_path}")
            return handler(request, session, document)
        finally:
            _close_document(document, save=False)


def _list_sheets(
    _request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    sheets = document.getSheets()
    names = [sheets.getByIndex(i).getName() for i in range(sheets.getCount())]
    return {"sheets": names, "count": len(names)}


def _read_cell(request: dict[str, Any], _session: dict[str, Any], document: Any) -> dict[str, Any]:
    sheet = _get_sheet(document, request["sheet_name"])
    cell = sheet.getCellRangeByName(str(request["cell_address"]))
    cell_type = _cell_type_name(cell.getType())
    value = None
    formula = None

    if cell_type == "TEXT":
        value = cell.getString()
    elif cell_type in ("VALUE", "FORMULA"):
        value = cell.getValue()
        if value == 0.0:
            cell_text = cell.getString()
            if cell_text and cell_text.strip():
                value = cell_text

    if cell_type == "FORMULA":
        formula = cell.getFormula()

    return {"value": _jsonable(value), "formula": formula, "type": cell_type}


def _get_sheet_data(
    request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    sheet = _get_sheet(document, request["sheet_name"])
    cell_range = sheet.getCellRangeByName(str(request["range_address"]))
    data = [list(row) for row in cell_range.getDataArray()]
    return {
        "data": _jsonable(data),
        "rows": len(data),
        "cols": len(data[0]) if data else 0,
    }


def _get_cell_colors(
    request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    sheet = _get_sheet(document, request["sheet_name"])
    cell_range = sheet.getCellRangeByName(str(request["range_address"]))
    colors = []
    for row_idx in range(cell_range.getRows().getCount()):
        row_colors = []
        for col_idx in range(cell_range.getColumns().getCount()):
            cell = cell_range.getCellByPosition(col_idx, row_idx)
            row_colors.append(cell.getPropertyValue("CellBackColor"))
        colors.append(row_colors)
    return {
        "colors": _jsonable(colors),
        "rows": len(colors),
        "cols": len(colors[0]) if colors else 0,
    }


def _inspect_document_cells(
    request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    """Read requested formula cells for deterministic host-side transformation."""
    document.calculateAll()
    sheets = document.getSheets()
    sheet_names = [sheets.getByIndex(index).getName() for index in range(sheets.getCount())]
    cells = []
    for item in request.get("cells") or []:
        sheet_name = str(item["sheet"])
        address = str(item["address"])
        if not sheets.hasByName(sheet_name):
            cells.append({"sheet": sheet_name, "address": address, "found": False, "error": None})
            continue
        cell = sheets.getByName(sheet_name).getCellRangeByName(address)
        cell_type = _cell_type_name(cell.getType())
        cells.append(
            {
                "sheet": sheet_name,
                "address": address,
                "found": True,
                "formula": cell.getFormula(),
                "error": cell.getError(),
                "type": cell_type,
                "value": _cell_value(cell, cell_type),
            }
        )
    return {"sheet_names": sheet_names, "cells": cells}


def _list_formula_cells(
    _request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    """Return every formula cell and its evaluated target-runtime state."""
    cells = []
    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        sheet = sheets.getByIndex(sheet_index)
        formula_ranges = sheet.queryContentCells(16).getRangeAddresses()
        for cell_range in formula_ranges:
            for row in range(cell_range.StartRow, cell_range.EndRow + 1):
                for column in range(cell_range.StartColumn, cell_range.EndColumn + 1):
                    cell = sheet.getCellByPosition(column, row)
                    cell_type = _cell_type_name(cell.getType())
                    cells.append(
                        {
                            "sheet": sheet.getName(),
                            "row": row,
                            "column": column,
                            "address": _a1_address(column, row),
                            "formula": cell.getFormula(),
                            "error": cell.getError(),
                            "type": cell_type,
                            "value": _cell_value(cell, cell_type),
                        }
                    )
    return {"cells": cells, "count": len(cells), "formula_count": len(cells)}


def _a1_address(column: int, row: int) -> str:
    """Return a zero-based Calc cell position as an A1 address."""
    letters = ""
    value = column + 1
    while value:
        value, remainder = divmod(value - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return f"{letters}{row + 1}"


def _cell_value(cell: Any, cell_type: str) -> Any:
    if cell_type == "TEXT":
        return cell.getString()
    value = cell.getValue()
    if value == 0.0:
        text = cell.getString()
        if text and text.strip():
            return text
    return _jsonable(value)


def _apply_document_repairs(request: dict[str, Any]) -> dict[str, Any]:
    """Apply explicit repairs and write a new ODS without mutating the staged source."""
    source = Path(str(request["ods_path"])).resolve()
    output = Path(str(request["output_path"])).resolve()
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            source_url = session["uno"].systemPathToFileUrl(str(source))
            document = session["desktop"].loadComponentFromURL(
                source_url, "_blank", 0, (_property_value("Hidden", True),)
            )
            if document is None:
                raise RuntimeError(f"LibreOffice could not open document: {source}")
            named_ranges = document.getPropertyValue("NamedRanges")
            address = _uno_struct("com.sun.star.table.CellAddress")
            named_ranges_added = 0
            for item in request.get("named_ranges") or []:
                name = str(item["name"])
                if named_ranges.hasByName(name):
                    continue
                named_ranges.addNewByName(name, str(item["content"]), address, 0)
                named_ranges_added += 1

            sheets = document.getSheets()
            formulas_applied = 0
            for item in request.get("formula_repairs") or []:
                sheet = sheets.getByName(str(item["sheet"]))
                cell = sheet.getCellRangeByName(str(item["address"]))
                cell.setFormula(str(item["formula"]))
                formulas_applied += 1
            document.calculateAll()
            output_url = session["uno"].systemPathToFileUrl(str(output))
            document.storeToURL(
                output_url,
                (
                    _property_value("FilterName", "calc8"),
                    _property_value("Overwrite", True),
                ),
            )
        finally:
            _close_document(document, save=False)
    if not output.is_file():
        raise RuntimeError("LibreOffice did not produce the repaired ODS output")
    return {
        "named_ranges_added": named_ranges_added,
        "formulas_applied": formulas_applied,
        "output_sha256": _sha256_file(output),
    }


def _execute_script(
    request: dict[str, Any], session: dict[str, Any], document: Any
) -> dict[str, Any]:
    script_uri = str(request["script_uri"])
    if _is_document_python_script(script_uri):
        return _execute_embedded_python_script(request, session, document)
    factory = session["component_context"].ServiceManager.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory",
        session["component_context"],
    )
    script_provider = factory.createScriptProvider(document)
    script = script_provider.getScript(script_uri)
    return_value = script.invoke((), (), ())
    return {
        "executed": True,
        "executor": "libreoffice-script-provider",
        "return_value": _jsonable(return_value),
    }


def _is_document_python_script(script_uri: str) -> bool:
    """Return whether the URI addresses a document-local Python macro."""
    if not script_uri.startswith("vnd.sun.star.script:"):
        return False
    _, _, raw_query = script_uri.partition("?")
    query = parse_qs(raw_query)
    return (
        query.get("language", [""])[0].casefold() == "python"
        and query.get("location", [""])[0].casefold() == "document"
    )


class _DocumentScriptContext:
    """Minimal XSCRIPTCONTEXT implementation backed by the active UNO session."""

    def __init__(self, document: Any, session: dict[str, Any]) -> None:
        self._document = document
        self._session = session

    def getDocument(self) -> Any:
        return self._document

    def getDesktop(self) -> Any:
        return self._session["desktop"]

    def getComponentContext(self) -> Any:
        return self._session["component_context"]


def _execute_embedded_python_script(
    request: dict[str, Any], session: dict[str, Any], document: Any
) -> dict[str, Any]:
    """Execute an embedded Python macro with the live UNO document context."""
    script_uri = str(request["script_uri"])
    identifier, separator, _query = script_uri.removeprefix("vnd.sun.star.script:").partition("?")
    if not separator:
        raise ValueError("document Python script URI has no query")
    raw_module, function_separator, raw_function = identifier.partition("$")
    module_name = unquote(raw_module)
    function_name = unquote(raw_function)
    if (
        not function_separator
        or not module_name.endswith(".py")
        or Path(module_name).name != module_name
        or not function_name.isidentifier()
    ):
        raise ValueError(f"invalid document Python script URI: {script_uri}")

    ods_path = Path(str(request["ods_path"])).resolve()
    member = f"Scripts/python/{module_name}"
    with ZipFile(ods_path) as archive:
        source = archive.read(member).decode("utf-8")
    namespace: dict[str, Any] = {
        "__file__": f"{ods_path}!/{member}",
        "__name__": f"xlsliberator_document_macro_{Path(module_name).stem}",
        "XSCRIPTCONTEXT": _DocumentScriptContext(document, session),
    }
    # Macro execution is explicitly capability-gated by the caller and runs only
    # in the disposable, networkless office sandbox with a read-only root.
    exec(compile(source, namespace["__file__"], "exec"), namespace)  # nosec B102  # noqa: S102
    function = namespace.get(function_name)
    if not callable(function):
        raise ValueError(f"embedded Python macro is not callable: {function_name}")
    arguments = tuple(request.get("arguments") or ())
    return_value = function(*arguments)
    return {
        "executed": True,
        "executor": "libreoffice-bundled-python",
        "module": member,
        "return_value": _jsonable(return_value),
    }


def _execute_button_handler(
    request: dict[str, Any], session: dict[str, Any], document: Any
) -> dict[str, Any]:
    button_name = str(request["button_name"])
    button_found = False
    script_uri = None

    sheets = document.getSheets()
    for sheet_index in range(sheets.getCount()):
        sheet = sheets.getByIndex(sheet_index)
        forms = sheet.getDrawPage().getForms()
        for form_index in range(forms.getCount()):
            form = forms.getByIndex(form_index)
            for control_index in range(form.getCount()):
                control = form.getByIndex(control_index)
                if getattr(control, "Name", None) == button_name:
                    button_found = True
                    script_uri = _script_uri_from_control(control)
                    break
            if button_found:
                break
        if button_found:
            break

    if not button_found:
        raise ValueError(f"Button '{button_name}' not found")

    if not script_uri:
        ods_path = request.get("ods_path")
        if ods_path:
            script_uri = _script_uri_from_content_xml(Path(str(ods_path)), button_name)

    if not script_uri:
        raise ValueError(f"Button '{button_name}' has no event handler")

    execution_data = _execute_script(
        {"script_uri": script_uri, "ods_path": request.get("ods_path")},
        session,
        document,
    )
    return {
        "handler_executed": True,
        "button_name": button_name,
        "script_uri": script_uri,
        "script_result": execution_data.get("return_value"),
    }


def _recalculate_document(
    _request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    document.calculateAll()
    return {"recalculated": True}


def _run_scenario(request: dict[str, Any]) -> dict[str, Any]:
    """Execute a versioned scenario against a disposable workbook copy."""
    source = Path(str(request["ods_path"])).resolve()
    if not source.is_file():
        raise FileNotFoundError(f"Scenario input not found: {source}")
    scenario = request.get("scenario")
    environment = request.get("environment")
    if not isinstance(scenario, dict) or not isinstance(environment, dict):
        raise ValueError("run_scenario requires scenario and environment objects")

    source_hash_before = _sha256_file(source)
    started_at = _utc_now()
    steps: list[dict[str, Any]] = []
    overall_status = "passed"
    final_path: Path | None = None
    session_runtime: dict[str, Any] = {}
    generated_attachments: list[Path] = []
    attachment_records: list[dict[str, str]] = []

    with tempfile.TemporaryDirectory(prefix="xlsliberator-scenario-") as workspace_name:
        workspace = Path(workspace_name)
        working_path = workspace / "working-copy.ods"
        shutil.copy2(source, working_path)
        current_path = working_path
        office_session = _office_session(request, use_gui=False)
        with office_session as session:
            session_runtime = {
                "profile_identifier": session["profile_identifier"],
                "profile_path": session["profile_path"],
                "pipe_name": session["pipe_name"],
                "office_executable": session["office_executable"],
                "office_pid": session["office_pid"],
            }
            document = None
            try:
                for raw_step in scenario.get("steps") or []:
                    if not isinstance(raw_step, dict):
                        raise ValueError("scenario steps must be objects")
                    step_started = _utc_now()
                    action = raw_step.get("action") or {}
                    action_kind = str(action.get("kind") or "")
                    required = bool(action.get("required", True))
                    before, before_errors = _scenario_observations(
                        raw_step.get("observations_before"),
                        document,
                        current_path,
                        environment,
                    )
                    action_status = "passed"
                    error: dict[str, Any] | None = None
                    evidence: list[str] = []
                    try:
                        document, current_path, action_evidence = _scenario_action(
                            action_kind,
                            dict(action.get("parameters") or {}),
                            document,
                            current_path,
                            workspace,
                            session,
                            environment,
                        )
                        evidence.extend(action_evidence)
                    except _ScenarioUnavailable as exc:
                        action_status = "unavailable"
                        error = {"type": type(exc).__name__, "message": str(exc)}
                    except Exception as exc:
                        action_status = "failed"
                        error = _error_payload(exc)

                    after, after_errors = _scenario_observations(
                        raw_step.get("observations_after"),
                        document,
                        current_path,
                        environment,
                    )
                    observation_errors = before_errors + after_errors
                    if observation_errors and action_status == "passed":
                        action_status = "failed"
                        error = {
                            "type": "ObservationError",
                            "message": "; ".join(observation_errors),
                        }
                    steps.append(
                        {
                            "step_id": str(raw_step.get("id") or ""),
                            "action": action_kind,
                            "status": action_status,
                            "started_at": step_started,
                            "ended_at": _utc_now(),
                            "observations_before": before,
                            "observations_after": after,
                            "evidence": evidence,
                            "error": error,
                        }
                    )
                    if required and action_status != "passed":
                        overall_status = action_status
                        break

                if document is not None and bool(request.get("final_save_reopen", True)):
                    final_path = workspace / "final-saved-reopened.ods"
                    _store_document(document, final_path, session)
                    _close_document(document, save=False)
                    document = _open_scenario_document(final_path, session)
                    evidence_step: dict[str, Any] = {
                        "step_id": "__final_save_close_reopen__",
                        "action": "reopen",
                        "status": "passed",
                        "started_at": _utc_now(),
                        "ended_at": _utc_now(),
                        "observations_before": {},
                        "observations_after": {},
                        "evidence": ["final save/close/reopen completed"],
                        "error": None,
                    }
                    steps.append(evidence_step)
                elif document is not None:
                    _close_document(document, save=False)
                    document = None
                    steps.append(
                        {
                            "step_id": "__final_close__",
                            "action": "close",
                            "status": "passed",
                            "started_at": _utc_now(),
                            "ended_at": _utc_now(),
                            "observations_before": {},
                            "observations_after": {},
                            "evidence": ["document closed without mutating the source fixture"],
                            "error": None,
                        }
                    )
            finally:
                generated_attachments = [
                    Path(str(item)) for item in session.get("generated_attachments", [])
                ]
                _close_document(document, save=False)

        session_runtime["office_exit_code"] = office_session.data.get("office_exit_code")
        logs = [str(office_session.data.get("office_log") or "")]
        final_working_copy = final_path or current_path
        final_hash = _sha256_file(final_working_copy)
        output_path = _job_output_path(request, "output_path")
        if output_path is not None:
            shutil.copy2(final_working_copy, output_path)
        attachment_output = _job_output_path(request, "attachment_output_path")
        if attachment_output is not None:
            if len(generated_attachments) != 1 or not generated_attachments[0].is_file():
                raise RuntimeError(
                    "scenario requested one attachment output but did not produce exactly one file"
                )
            shutil.copy2(generated_attachments[0], attachment_output)
        attachment_records = [
            {"name": path.name, "sha256": _sha256_file(path)}
            for path in generated_attachments
            if path.is_file()
        ]

    source_hash_after = _sha256_file(source)
    if source_hash_after != source_hash_before:
        overall_status = "failed"
        steps.append(
            {
                "step_id": "__source_immutability__",
                "action": "close",
                "status": "failed",
                "started_at": _utc_now(),
                "ended_at": _utc_now(),
                "observations_before": {},
                "observations_after": {},
                "evidence": [],
                "error": {
                    "type": "SourceMutationError",
                    "message": "scenario execution mutated the staged source document",
                },
            }
        )
    return {
        "scenario_id": str(scenario.get("id") or ""),
        "status": overall_status,
        "started_at": started_at,
        "ended_at": _utc_now(),
        "workbook_hash_before": source_hash_before,
        "workbook_hash_after": source_hash_after,
        "final_working_copy_sha256": final_hash,
        "steps": steps,
        "runtime": session_runtime,
        "logs": logs,
        "attachments": attachment_records,
        "source_mutated": source_hash_before != source_hash_after,
    }


class _ScenarioUnavailable(RuntimeError):
    """Raised for an explicitly unsupported or unavailable scenario capability."""


def _scenario_action(
    kind: str,
    parameters: dict[str, Any],
    document: Any,
    current_path: Path,
    workspace: Path,
    session: dict[str, Any],
    environment: dict[str, Any],
) -> tuple[Any, Path, list[str]]:
    if kind == "execute_python_macro":
        kind = "invoke_macro"
    elif kind == "dispatch_control_event":
        raise _ScenarioUnavailable(
            "control dispatch requires the verified UI/event session layer from Prompt 06"
        )
    elif kind == "send_keyboard_event":
        raise _ScenarioUnavailable(
            "keyboard dispatch requires the verified UI/event session layer from Prompt 06"
        )
    elif kind == "export_pdf":
        kind = "export"
        parameters = {**parameters, "format": "pdf"}
    if kind == "open":
        if document is not None:
            raise RuntimeError("document is already open")
        document = _open_scenario_document(current_path, session)
        _configure_scenario_calculation(document, environment)
        return document, current_path, ["document opened with declared calculation settings"]
    if kind == "close":
        if document is None:
            raise RuntimeError("document is not open")
        _close_document(document, save=False)
        return None, current_path, ["document closed"]
    if kind == "reopen":
        _close_document(document, save=False)
        document = _open_scenario_document(current_path, session)
        _configure_scenario_calculation(document, environment)
        return document, current_path, ["document reopened with declared calculation settings"]
    if document is None:
        raise RuntimeError(f"action {kind} requires an open document")
    if kind == "set_cell":
        cell = _scenario_cell(document, parameters)
        if "formula" in parameters:
            cell.setFormula(str(parameters["formula"]))
        else:
            _set_scenario_cell_value(cell, parameters.get("value"))
        return document, current_path, ["cell updated"]
    if kind == "set_range":
        sheet = _get_sheet(document, parameters.get("sheet", parameters.get("sheet_name", 0)))
        address = str(parameters.get("range") or parameters.get("address") or "")
        values = parameters.get("values")
        if not address or not isinstance(values, list):
            raise ValueError("set_range requires range/address and two-dimensional values")
        rows = tuple(tuple(row) for row in values)
        if not rows or any(len(row) != len(rows[0]) for row in rows):
            raise ValueError("set_range values must be a non-empty rectangular matrix")
        target = sheet.getCellRangeByName(address)
        if target.getRows().getCount() != len(rows) or target.getColumns().getCount() != len(
            rows[0]
        ):
            raise ValueError("set_range values do not match the target range dimensions")
        for row_index, row in enumerate(rows):
            for column_index, value in enumerate(row):
                _set_scenario_cell_value(target.getCellByPosition(column_index, row_index), value)
        return document, current_path, ["range updated"]
    if kind == "recalculate":
        document.calculateAll()
        return document, current_path, ["calculateAll completed"]
    if kind == "activate_sheet":
        sheet = _get_sheet(document, parameters.get("sheet", parameters.get("sheet_name", 0)))
        document.getCurrentController().setActiveSheet(sheet)
        return document, current_path, [f"activated sheet {sheet.getName()}"]
    if kind == "copy_sheet":
        sheets = document.getSheets()
        source = str(parameters.get("source") or parameters.get("sheet") or "")
        target = str(parameters.get("target") or parameters.get("new_name") or "")
        index = int(parameters.get("index", sheets.getCount()))
        if not source or not target:
            raise ValueError("copy_sheet requires source and target")
        if sheets.hasByName(target):
            raise ValueError(f"copy_sheet target already exists: {target}")
        sheets.insertNewByName(target, index)
        try:
            source_sheet = sheets.getByName(source)
            target_sheet = sheets.getByName(target)
            cursor = source_sheet.createCursor()
            cursor.gotoEndOfUsedArea(True)
            target_sheet.copyRange(
                target_sheet.getCellByPosition(0, 0).getCellAddress(),
                cursor.getRangeAddress(),
            )
        except Exception:
            with suppress(Exception):
                sheets.removeByName(target)
            raise
        evidence = (
            f"copied used range from {source} to {target} at {index}; "
            "deterministic Calc compatibility copy used"
        )
        return document, current_path, [evidence]
    if kind == "move_sheet":
        sheets = document.getSheets()
        name = str(parameters.get("sheet") or parameters.get("name") or "")
        index = int(parameters.get("index", 0))
        if not name:
            raise ValueError("move_sheet requires sheet/name")
        session["uno"].invoke(
            sheets,
            "moveByName",
            (name, session["uno"].Any("short", index)),
        )
        return document, current_path, [f"moved sheet {name} to {index}"]
    if kind == "rename_sheet":
        sheet = _get_sheet(document, parameters.get("sheet", parameters.get("source", 0)))
        target = str(parameters.get("target") or parameters.get("new_name") or "")
        if not target:
            raise ValueError("rename_sheet requires target/new_name")
        source = sheet.getName()
        sheet.Name = target
        if sheet.getName() != target:
            persisted_path = workspace / f"before-rename-{uuid.uuid4().hex}.ods"
            _store_document(document, persisted_path, session)
            _close_document(document, save=False)
            current_path = persisted_path
            _rename_sheet_in_ods(current_path, source, target)
            document = _open_scenario_document(current_path, session)
            if not document.getSheets().hasByName(target):
                raise RuntimeError(f"LibreOffice did not reopen renamed sheet {target}")
            evidence = (
                f"renamed sheet {source} to {target}; deterministic ODF compatibility rename used"
            )
        else:
            evidence = f"renamed sheet {source} to {target}"
        return document, current_path, [evidence]
    if kind in {"save", "save_as"}:
        suffix = Path(str(parameters.get("filename") or "saved.ods")).suffix or ".ods"
        saved_path = workspace / f"scenario-{uuid.uuid4().hex}{suffix}"
        _store_document(document, saved_path, session, filter_name=_scenario_filter(suffix))
        return document, saved_path, [f"saved separate working copy: {saved_path.name}"]
    if kind == "invoke_macro":
        granted = _environment_grants(environment)
        if "macro_execution" not in granted:
            raise _ScenarioUnavailable("macro_execution capability was not granted")
        script_uri = str(parameters.get("script_uri") or "")
        if not script_uri:
            raise _ScenarioUnavailable("macro action requires a script_uri")
        try:
            result = _execute_script(
                {"script_uri": script_uri, "ods_path": str(current_path)},
                session,
                document,
            )
        except Exception as exc:
            raise _ScenarioUnavailable(
                f"LibreOffice scripting provider unavailable: {exc}"
            ) from exc
        return (
            document,
            current_path,
            [
                f"macro invoked: {script_uri}",
                f"macro executor: {result.get('executor')}",
                f"macro return value: {_jsonable(result.get('return_value'))}",
            ],
        )
    if kind == "click_control":
        granted = _environment_grants(environment)
        if "macro_execution" not in granted:
            raise _ScenarioUnavailable("click_control requires macro_execution capability")
        button_name = str(parameters.get("control_name") or parameters.get("button_name") or "")
        if not button_name:
            raise ValueError("click_control requires control_name/button_name")
        try:
            result = _execute_button_handler(
                {"button_name": button_name, "ods_path": str(current_path)},
                session,
                document,
            )
        except Exception as exc:
            raise _ScenarioUnavailable(f"control event execution unavailable: {exc}") from exc
        return (
            document,
            current_path,
            [
                f"control event invoked: {button_name}",
                f"script URI: {result.get('script_uri')}",
                f"script return value: {_jsonable(result.get('script_result'))}",
            ],
        )
    if kind == "refresh_data":
        refreshed = _refresh_scenario_data(document)
        return document, current_path, [f"refreshed data sources: {refreshed}"]
    if kind == "print":
        printed_path = workspace / f"printed-{uuid.uuid4().hex}.prn"
        try:
            document.print(
                (
                    _property_value(
                        "FileName", session["uno"].systemPathToFileUrl(str(printed_path))
                    ),
                    _property_value("Wait", True),
                )
            )
        except Exception as exc:
            raise _ScenarioUnavailable(f"LibreOffice print service unavailable: {exc}") from exc
        if not printed_path.is_file():
            raise RuntimeError("LibreOffice print action produced no output")
        return (
            document,
            current_path,
            [
                f"printed output: {printed_path.name}",
                f"printed sha256: {_sha256_file(printed_path)}",
            ],
        )
    if kind == "export":
        suffix = str(parameters.get("format") or parameters.get("suffix") or ".pdf")
        if not suffix.startswith("."):
            suffix = f".{suffix}"
        exported_path = workspace / f"export-{uuid.uuid4().hex}{suffix.lower()}"
        try:
            document.storeToURL(
                session["uno"].systemPathToFileUrl(str(exported_path)),
                (
                    _property_value("FilterName", _scenario_filter(suffix)),
                    _property_value("Overwrite", True),
                ),
            )
        except Exception as exc:
            raise _ScenarioUnavailable(
                f"LibreOffice export unavailable for {suffix}: {exc}"
            ) from exc
        if not exported_path.is_file():
            raise RuntimeError("LibreOffice export action produced no output")
        session.setdefault("generated_attachments", []).append(str(exported_path))
        return (
            document,
            current_path,
            [
                f"exported output: {exported_path.name}",
                f"exported sha256: {_sha256_file(exported_path)}",
            ],
        )
    raise _ScenarioUnavailable(f"unknown LibreOffice scenario action: {kind}")


def _environment_grants(environment: dict[str, Any]) -> set[str]:
    """Return both legacy and typed capability grants."""

    grants = {str(item) for item in environment.get("granted_capabilities") or []}
    for item in environment.get("typed_capabilities") or []:
        if isinstance(item, dict) and item.get("granted") and item.get("capability"):
            grants.add(str(item["capability"]))
    return grants


def _open_scenario_document(path: Path, session: dict[str, Any]) -> Any:
    document = session["desktop"].loadComponentFromURL(
        session["uno"].systemPathToFileUrl(str(path)),
        "_blank",
        0,
        (
            _property_value("MacroExecutionMode", 4),
            _property_value("Hidden", True),
        ),
    )
    if document is None:
        raise RuntimeError(f"LibreOffice could not open scenario document: {path}")
    return document


def _store_document(
    document: Any,
    path: Path,
    session: dict[str, Any],
    *,
    filter_name: str = "calc8",
) -> None:
    document.storeAsURL(
        session["uno"].systemPathToFileUrl(str(path)),
        (
            _property_value("FilterName", filter_name),
            _property_value("Overwrite", True),
        ),
    )
    if not path.is_file():
        raise RuntimeError(f"LibreOffice did not produce scenario output: {path}")


def _scenario_filter(suffix: str) -> str:
    normalized = suffix.lower().lstrip(".")
    filters = {
        "ods": "calc8",
        "xlsx": "Calc MS Excel 2007 XML",
        "xls": "MS Excel 97",
        "csv": "Text - txt - csv (StarCalc)",
        "pdf": "calc_pdf_Export",
    }
    try:
        return filters[normalized]
    except KeyError as exc:
        raise _ScenarioUnavailable(f"unsupported LibreOffice export format: {suffix}") from exc


def _refresh_scenario_data(document: Any) -> int:
    """Refresh every Calc data source, treating an empty set as a successful no-op."""
    with suppress(AttributeError):
        document.refresh()
        return 1
    refreshed = 0
    for property_name in ("DatabaseRanges", "AreaLinks", "DDELinks", "ExternalDocLinks"):
        with suppress(Exception):
            collection = document.getPropertyValue(property_name)
            names = collection.getElementNames()
            for name in names:
                item = collection.getByName(name)
                refresh = getattr(item, "refresh", None)
                if callable(refresh):
                    refresh()
                    refreshed += 1
    return refreshed


def _scenario_observations(
    requests: Any,
    document: Any,
    current_path: Path,
    environment: dict[str, Any],
) -> tuple[dict[str, Any], list[str]]:
    values: dict[str, Any] = {}
    errors: list[str] = []
    for raw in requests or []:
        observation_id = str(raw.get("id") or "")
        try:
            values[observation_id] = _scenario_observation(
                str(raw.get("kind") or ""),
                dict(raw.get("selector") or {}),
                document,
                current_path,
                environment,
            )
        except Exception as exc:
            values[observation_id] = {
                "kind": "error",
                "value": None,
                "error_type": type(exc).__name__,
                "metadata": {"message": str(exc)},
            }
            if bool(raw.get("required", True)):
                errors.append(f"{observation_id}: {exc}")
    return values, errors


def _scenario_observation(
    kind: str,
    selector: dict[str, Any],
    document: Any,
    current_path: Path,
    environment: dict[str, Any],
) -> dict[str, Any]:
    if kind in {"cell", "cell_value"}:
        if document is None:
            raise RuntimeError("cell observation requires an open document")
        cell = _scenario_cell(document, selector)
        cell_type = _cell_type_name(cell.getType())
        error_code = int(cell.getError())
        formula = cell.getFormula() if cell_type == "FORMULA" else None
        if error_code:
            display = cell.getString()
            error_type = display if display.startswith("#") else f"libreoffice:{error_code}"
            return {
                "kind": "error",
                "value": display or error_code,
                "error_type": error_type,
                "formula": formula,
                "cell_type": cell_type,
                "metadata": {"libreoffice_error_code": error_code},
            }
        return _scenario_normalized_value(
            _scenario_cell_value(document, cell, cell_type) if cell_type != "EMPTY" else None,
            environment,
            formula=formula,
            cell_type=cell_type,
        )
    if kind in {"cell_formula", "cell_type", "cell_error"}:
        if document is None:
            raise RuntimeError(f"{kind} observation requires an open document")
        cell = _scenario_cell(document, selector)
        cell_type = _cell_type_name(cell.getType())
        error_code = int(cell.getError())
        if kind == "cell_formula":
            formula = cell.getFormula() if cell_type == "FORMULA" else ""
            return _scenario_normalized_value(formula, environment, cell_type=cell_type)
        if kind == "cell_type":
            return {"kind": "string", "value": cell_type, "cell_type": cell_type}
        if not error_code:
            return {"kind": "empty_cell", "value": None, "cell_type": cell_type}
        display = cell.getString()
        return {
            "kind": "error",
            "value": display or error_code,
            "error_type": display if display.startswith("#") else f"libreoffice:{error_code}",
            "cell_type": cell_type,
            "metadata": {"libreoffice_error_code": error_code},
        }
    if kind == "range_values":
        if document is None:
            raise RuntimeError("range observation requires an open document")
        target = _scenario_cell(document, selector)
        rows = target.getRows().getCount()
        columns = target.getColumns().getCount()
        values = [
            [
                _scenario_normalized_value(
                    _scenario_cell_value(
                        document,
                        target.getCellByPosition(column, row),
                        _cell_type_name(target.getCellByPosition(column, row).getType()),
                    )
                    if _cell_type_name(target.getCellByPosition(column, row).getType()) != "EMPTY"
                    else None,
                    environment,
                    cell_type=_cell_type_name(target.getCellByPosition(column, row).getType()),
                )
                for column in range(columns)
            ]
            for row in range(rows)
        ]
        return {"kind": "array", "value": values}
    if kind in {"sheets", "sheet_state"}:
        if document is None:
            raise RuntimeError("sheet observation requires an open document")
        sheets = document.getSheets()
        inventory = []
        for index in range(sheets.getCount()):
            sheet = sheets.getByIndex(index)
            inventory.append(
                {"index": index, "name": sheet.getName(), "visible": bool(sheet.IsVisible)}
            )
        return {"kind": "object", "value": {"sheets": inventory}}
    if kind == "named_ranges":
        if document is None:
            raise RuntimeError("named-range observation requires an open document")
        names = document.getPropertyValue("NamedRanges")
        inventory = []
        for name in sorted(names.getElementNames()):
            item = names.getByName(name)
            inventory.append({"name": name, "content": item.getContent()})
        return {"kind": "object", "value": {"named_ranges": inventory}}
    if kind == "embedded_scripts":
        return {"kind": "object", "value": _scenario_script_inventory(current_path)}
    if kind in {"controls_events", "controls_bindings"}:
        return {"kind": "object", "value": _scenario_control_event_inventory(current_path)}
    if kind == "package_hash":
        return {"kind": "string", "value": _sha256_file(current_path)}
    if kind == "artifact_inventory":
        return {"kind": "object", "value": _scenario_package_inventory(current_path)}
    if kind == "runtime_errors":
        return {"kind": "array", "value": []}
    if kind in {"files_created", "mocked_calls", "screenshots"}:
        raise _ScenarioUnavailable(
            f"{kind} observation requires the stateful runtime evidence layer from Prompt 06"
        )
    raise _ScenarioUnavailable(f"unknown LibreOffice observation kind: {kind}")


def _scenario_cell(document: Any, selector: dict[str, Any]) -> Any:
    sheet = _get_sheet(document, selector.get("sheet", selector.get("sheet_name", 0)))
    address = str(selector.get("address") or selector.get("cell_address") or "")
    if not address:
        raise ValueError("cell selector requires address")
    return sheet.getCellRangeByName(address)


def _set_scenario_cell_value(cell: Any, value: Any) -> None:
    if value is None:
        cell.setString("")
    elif isinstance(value, bool):
        cell.setValue(1.0 if value else 0.0)
    elif isinstance(value, (int, float)):
        cell.setValue(float(value))
    else:
        cell.setString(str(value))


def _configure_scenario_calculation(document: Any, environment: dict[str, Any]) -> None:
    """Apply calculation mode and iteration settings, failing if requested semantics are absent."""
    mode = str(environment.get("calculation_mode") or "automatic")
    if mode == "automatic_except_tables":
        raise _ScenarioUnavailable(
            "LibreOffice has no verified automatic-except-data-tables calculation mode"
        )
    document.enableAutomaticCalculation(mode == "automatic")
    requested_iteration = bool(environment.get("iterative_calculation", False))
    settings = {
        "IsIterationEnabled": requested_iteration,
        "IterationCount": int(environment.get("max_iterations") or 100),
        "IterationEpsilon": float(environment.get("max_change") or 0.001),
    }
    for name, value in settings.items():
        try:
            document.setPropertyValue(name, value)
        except Exception as exc:
            if requested_iteration:
                raise _ScenarioUnavailable(
                    f"LibreOffice cannot apply iterative calculation property {name}: {exc}"
                ) from exc


def _scenario_cell_value(document: Any, cell: Any, cell_type: str) -> Any:
    """Return typed target values, including dates using Calc's actual null date."""
    value = _cell_value(cell, cell_type)
    if cell_type not in {"VALUE", "FORMULA"} or not isinstance(value, (int, float)):
        return value
    try:
        format_key = int(cell.getPropertyValue("NumberFormat"))
        format_type = int(document.getNumberFormats().getByKey(format_key).getPropertyValue("Type"))
        is_date = bool(format_type & 2)
        has_time = bool(format_type & 4)
        if not is_date:
            return value
        null_date = document.getPropertyValue("NullDate")
        base = datetime(int(null_date.Year), int(null_date.Month), int(null_date.Day))
        converted = base + timedelta(days=float(value))
        return converted if has_time else converted.date()
    except Exception:
        return value


def _scenario_normalized_value(
    value: Any,
    environment: dict[str, Any],
    *,
    formula: str | None = None,
    cell_type: str | None = None,
) -> dict[str, Any]:
    common = {"formula": formula, "cell_type": cell_type}
    if value is None:
        return {"kind": "empty_cell", "value": None, **common}
    if isinstance(value, bool):
        return {"kind": "boolean", "value": value, **common}
    if isinstance(value, datetime):
        return {
            "kind": "datetime",
            "value": value.isoformat(),
            "date_system": str(environment.get("date_system") or "1900"),
            "timezone": str(environment.get("timezone") or "UTC"),
            **common,
        }
    if isinstance(value, date):
        return {
            "kind": "date",
            "value": value.isoformat(),
            "date_system": str(environment.get("date_system") or "1900"),
            "timezone": str(environment.get("timezone") or "UTC"),
            **common,
        }
    if isinstance(value, (int, float)):
        return {"kind": "number", "value": value, **common}
    if isinstance(value, str):
        return {"kind": "empty_string" if value == "" else "string", "value": value, **common}
    return {
        "kind": "object",
        "value": _jsonable(value),
        "metadata": {
            "date_system": str(environment.get("date_system") or "1900"),
            "timezone": str(environment.get("timezone") or "UTC"),
        },
        **common,
    }


def _scenario_script_inventory(path: Path) -> dict[str, Any]:
    with ZipFile(path) as archive:
        names = sorted(
            name
            for name in archive.namelist()
            if name.startswith(("Scripts/", "Basic/")) or "script" in name.lower()
        )
    return {"entries": names, "count": len(names)}


def _scenario_control_event_inventory(path: Path) -> dict[str, Any]:
    # This worker must remain standard-library-only for LibreOffice's bundled Python.
    import xml.etree.ElementTree as ET  # nosec B405

    with ZipFile(path) as archive:
        payload = archive.read("content.xml")
    normalized_payload = payload.upper()
    if b"<!DOCTYPE" in normalized_payload or b"<!ENTITY" in normalized_payload:
        raise ValueError("content.xml must not contain DTD or entity declarations")
    # Parsing is safe here because declarations were rejected before ElementTree sees the bytes.
    root = ET.fromstring(payload)  # nosec B314
    controls = []
    events = []
    for element in root.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name in {"button", "checkbox", "combobox", "listbox", "form"}:
            controls.append(
                {"type": local_name, "attributes": dict(sorted(element.attrib.items()))}
            )
        if "event" in local_name:
            events.append({"type": local_name, "attributes": dict(sorted(element.attrib.items()))})
    return {"controls": controls, "events": events}


def _scenario_package_inventory(path: Path) -> dict[str, Any]:
    with ZipFile(path) as archive:
        parts = []
        for info in sorted(archive.infolist(), key=lambda item: item.filename):
            parts.append(
                {
                    "name": info.filename,
                    "size": info.file_size,
                    "sha256": hashlib.sha256(archive.read(info.filename)).hexdigest(),
                }
            )
    return {"parts": parts, "count": len(parts)}


def _rename_sheet_in_ods(path: Path, source: str, target: str) -> None:
    """Apply a package-level sheet rename when the stock UNO setter is ineffective."""
    source_xml = _escape_xml_attribute(source)
    target_xml = _escape_xml_attribute(target)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        with ZipFile(path, "r") as archive, ZipFile(temporary, "w") as output:
            renamed = False
            for item in archive.infolist():
                payload = archive.read(item.filename)
                if item.filename == "content.xml":
                    text = payload.decode("utf-8")
                    old_declaration = f'table:name="{source_xml}"'
                    new_declaration = f'table:name="{target_xml}"'
                    if old_declaration not in text:
                        raise ValueError(f"sheet declaration not found in ODS: {source}")
                    text = text.replace(old_declaration, new_declaration, 1)
                    text = text.replace(f"${source_xml}.", f"${target_xml}.")
                    payload = text.encode("utf-8")
                    renamed = True
                output.writestr(item, payload)
            if not renamed:
                raise ValueError("ODS package has no content.xml")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _escape_xml_attribute(value: str) -> str:
    """Escape text for a double-quoted XML attribute without parsing XML."""
    return (
        value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")
    )


def _utc_now() -> str:
    return datetime.now(UTC).isoformat()


def _job_output_path(request: dict[str, Any], key: str) -> Path | None:
    raw = request.get(key)
    if not raw:
        return None
    path = Path(str(raw))
    resolved = path.resolve()
    if not resolved.is_relative_to(Path("/job")) or path.is_symlink():
        raise ValueError(f"{key} must be a regular output path beneath /job")
    if not resolved.parent.is_dir():
        raise FileNotFoundError(f"{key} parent does not exist: {resolved.parent}")
    return resolved


class _OfficeSession:
    def __init__(self, request: dict[str, Any], use_gui: bool) -> None:
        self.request = request
        self.use_gui = use_gui
        self.process: subprocess.Popen[bytes] | None = None
        self.tmpdir: tempfile.TemporaryDirectory[str] | None = None
        self.log_handle: Any = None
        self.log_path: Path | None = None
        self.data: dict[str, Any] = {}

    def __enter__(self) -> dict[str, Any]:
        import uno

        executable = _authorized_office_executable()
        requested_profile = str(self.request.get("session_profile_identifier") or "")
        if requested_profile and not re.fullmatch(r"[A-Za-z0-9_-]{1,80}", requested_profile):
            raise ValueError("session profile identifier is malformed")
        profile_identifier = requested_profile or f"profile-{uuid.uuid4().hex}"
        requested_port = self.request.get("session_port")
        port = int(requested_port) if requested_port is not None else None
        if port is not None and not 1024 <= port <= 65535:
            raise ValueError("session UNO port must be between 1024 and 65535")
        display = str(self.request.get("session_display") or "")
        if display and not re.fullmatch(r":\d{1,5}", display):
            raise ValueError("session display identifier is malformed")
        pipe_name = f"xlsliberator_{profile_identifier}"
        self.tmpdir = tempfile.TemporaryDirectory(prefix="xlsliberator-lo-worker-")
        profile_dir = Path(self.tmpdir.name) / profile_identifier
        profile_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = Path(self.tmpdir.name) / "office.log"
        self.log_handle = self.log_path.open("wb")

        cmd = [
            executable,
            "--nologo",
            "--nodefault",
            "--norestore",
            "--nofirststartwizard",
            "--nolockcheck",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
        ]
        if port is None:
            cmd.append(f"--accept=pipe,name={pipe_name};urp;")
            uno_url = f"uno:pipe,name={pipe_name};urp;StarOffice.ComponentContext"
        else:
            cmd.append(f"--accept=socket,host=127.0.0.1,port={port};urp;")
            uno_url = f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"
        if not self.use_gui:
            cmd.insert(1, "--headless")

        env = dict(os.environ)
        env["SAL_USE_VCLPLUGIN"] = "svp"
        if display:
            env["DISPLAY"] = display
        self.process = subprocess.Popen(
            cmd,
            env=env,
            stdout=self.log_handle,
            stderr=subprocess.STDOUT,
        )
        try:
            local_context = uno.getComponentContext()
            local_service_manager = local_context.getServiceManager()
            resolver = local_service_manager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context
            )
            component_context = _resolve_component_context(
                resolver,
                uno_url,
                self.process,
                int(self.request.get("start_timeout_seconds", 20)),
            )
            remote_service_manager = component_context.getServiceManager()
            desktop = remote_service_manager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", component_context
            )
        except Exception as exc:
            raise RuntimeError(self._startup_failure_message(exc)) from exc
        self.data = {
            "uno": uno,
            "component_context": component_context,
            "desktop": desktop,
            "pipe_name": pipe_name,
            "profile_identifier": profile_identifier,
            "profile_path": str(profile_dir),
            "uno_port": port,
            "display": display or None,
            "office_executable": executable,
            "office_pid": self.process.pid,
        }
        return self.data

    def __exit__(self, _exc_type: Any, _exc: Any, _tb: Any) -> None:
        if self.process is not None:
            self.process.terminate()
            try:
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
                self.process.wait(timeout=10)
            self.data["office_exit_code"] = self.process.returncode
            self.process = None
        if self.log_handle is not None:
            self.log_handle.flush()
            self.log_handle.close()
            self.log_handle = None
        if self.log_path is not None and self.log_path.is_file():
            self.data["office_log"] = self.log_path.read_text(errors="replace")[-16000:]
        if self.tmpdir is not None:
            self.tmpdir.cleanup()
            self.tmpdir = None

    def _startup_failure_message(self, exc: BaseException) -> str:
        """Attach bounded container-local office logs to startup failures."""
        time.sleep(0.1)
        exit_code = self.process.poll() if self.process is not None else None
        if self.log_handle is not None:
            self.log_handle.flush()
        log = ""
        if self.log_path is not None and self.log_path.is_file():
            log = self.log_path.read_text(errors="replace")[-4000:].strip()
        return f"LibreOffice startup failed: {exc}; exit_code={exit_code}; log={log or '<empty>'}"


def _office_session(request: dict[str, Any], use_gui: bool) -> _OfficeSession:
    return _OfficeSession(request, use_gui)


def _property_value(name: str, value: Any) -> Any:
    prop = _uno_struct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def _uno_struct(type_name: str) -> Any:
    """Create a UNO struct without relying on generated ``com.sun.star`` modules."""
    uno = importlib.import_module("uno")
    return uno.createUnoStruct(type_name)


def _set_parser_property(parser: Any, name: str, value: Any) -> None:
    with suppress(Exception):
        parser.setPropertyValue(name, value)


def _formula_token_to_string(token: Any) -> str:
    opcode = getattr(token, "OpCode", None)
    data = getattr(token, "Data", None)
    if data is None or data == "":
        return str(opcode)
    return f"{opcode}:{data}"


def _formula_lexical_errors(formula: str) -> list[str]:
    """Reject structural syntax that LibreOffice's parser silently repairs.

    This is a supplementary fail-closed check. Successful certification still
    requires the target FormulaParser and an equivalent token round-trip.
    """
    pairs = {")": "(", "]": "[", "}": "{"}
    opening = set(pairs.values())
    stack: list[tuple[str, int]] = []
    errors: list[str] = []
    in_string = False
    index = 0
    while index < len(formula):
        character = formula[index]
        if character == '"':
            if in_string and index + 1 < len(formula) and formula[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string and character in opening:
            stack.append((character, index))
        elif not in_string and character in pairs:
            if not stack:
                errors.append(f"unexpected {character!r} at offset {index}")
            else:
                opened, opened_at = stack.pop()
                if opened != pairs[character]:
                    errors.append(
                        f"mismatched {opened!r} at offset {opened_at} and "
                        f"{character!r} at offset {index}"
                    )
        index += 1
    if in_string:
        errors.append("unterminated string literal")
    for opened, opened_at in reversed(stack):
        errors.append(f"unclosed {opened!r} at offset {opened_at}")
    return errors


def _get_sheet(document: Any, name_or_index: Any) -> Any:
    sheets = document.getSheets()
    if isinstance(name_or_index, str) and name_or_index.isdigit():
        name_or_index = int(name_or_index)
    if isinstance(name_or_index, int):
        if name_or_index < 0 or name_or_index >= sheets.getCount():
            raise IndexError(f"Sheet index out of range: {name_or_index}")
        return sheets.getByIndex(name_or_index)
    if not sheets.hasByName(str(name_or_index)):
        raise KeyError(f"Sheet not found: {name_or_index}")
    return sheets.getByName(str(name_or_index))


def _cell_type_name(cell_type: Any) -> str:
    value = getattr(cell_type, "value", cell_type)
    if isinstance(value, str):
        return value
    type_map = {0: "EMPTY", 1: "VALUE", 2: "TEXT", 3: "FORMULA"}
    with suppress(Exception):
        return type_map[int(value)]
    text = str(cell_type)
    for name in ("EMPTY", "VALUE", "TEXT", "FORMULA"):
        if name in text:
            return name
    return text


def _script_uri_from_control(control: Any) -> str | None:
    try:
        events = getattr(control, "Events", None)
        if events and events.hasByName("approveAction"):
            event = events.getByName("approveAction")
            for prop in event:
                if prop.Name == "Script":
                    return str(prop.Value)
    except Exception:
        return None
    return None


def _script_uri_from_content_xml(ods_path: Path, button_name: str) -> str | None:
    import re

    with ZipFile(ods_path, "r") as archive:
        content = archive.read("content.xml").decode("utf-8")
    pattern = rf'form:name="{re.escape(button_name)}"[^>]*>.*?xlink:href="([^"]+)"'
    match = re.search(pattern, content, re.DOTALL)
    if not match:
        return None
    return match.group(1).replace("&amp;", "&")


def _close_document(document: Any, save: bool) -> None:
    if document is None:
        return
    try:
        document.close(save)
    except Exception:
        with suppress(Exception):
            document.dispose()


def _resolve_component_context(
    resolver: Any,
    uno_url: str,
    process: subprocess.Popen[bytes],
    timeout_seconds: int,
) -> Any:
    """Wait for readiness by completing a real UNO handshake."""
    deadline = time.monotonic() + timeout_seconds
    last_error: BaseException | None = None
    while time.monotonic() < deadline:
        exit_code = process.poll()
        if exit_code is not None:
            raise RuntimeError(f"LibreOffice exited before UNO readiness with code {exit_code}")
        try:
            return resolver.resolve(uno_url)
        except Exception as exc:
            last_error = exc
        time.sleep(0.2)
    raise TimeoutError(f"LibreOffice UNO endpoint did not become ready: {last_error}")


def _jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (list, tuple)):
        return [_jsonable(item) for item in value]
    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}
    return str(value)


def _error_payload(exc: BaseException) -> dict[str, str]:
    return {
        "type": type(exc).__name__,
        "message": str(exc),
        "traceback": traceback.format_exc(),
    }


def _write_response(response: dict[str, Any]) -> None:
    sys.stdout.write(json.dumps(_jsonable(response), ensure_ascii=False))
    sys.stdout.write("\n")
    sys.stdout.flush()


if __name__ == "__main__":
    raise SystemExit(main())
