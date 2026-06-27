"""LibreOffice Python worker for UNO operations.

This module is intentionally standard-library only so it can run under
LibreOffice's bundled Python wrapper.
"""

from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import traceback
from collections.abc import Callable
from contextlib import suppress
from pathlib import Path
from typing import Any
from zipfile import ZipFile

DEFAULT_START_TIMEOUT_SECONDS = 20


def main() -> int:
    """Read one JSON request from stdin and write one JSON response to stdout."""
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
    op = str(request.get("op", ""))
    if op == "ping":
        import uno

        return {
            "uno_importable": True,
            "uno_module": getattr(uno, "__file__", None),
            "python_executable": sys.executable,
        }
    if op == "parse_formula":
        return _parse_formula(request)
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
    if op == "click_form_button":
        return _with_document(request, _click_form_button)
    if op == "recalculate_document":
        return _with_document(request, _recalculate_document)
    raise ValueError(f"Unsupported worker op: {op}")


def _parse_formula(request: dict[str, Any]) -> dict[str, Any]:
    from com.sun.star.table import CellAddress

    formula = str(request["formula"])
    with _office_session(request, use_gui=False) as session:
        document = None
        try:
            desktop = session["desktop"]
            hidden = _property_value("Hidden", True)
            document = desktop.loadComponentFromURL(
                "private:factory/scalc", "_blank", 0, (hidden,)
            )
            if document is None:
                raise RuntimeError("LibreOffice did not create a Calc document")

            parser = document.createInstance("com.sun.star.sheet.FormulaParser")
            _set_parser_property(parser, "CompileEnglish", True)
            _set_parser_property(parser, "ParameterSeparator", ";")

            tokens = parser.parseFormula(formula, CellAddress(0, 0, 0))
            printed = parser.printFormula(tokens, CellAddress(0, 0, 0))
            return {
                "formula": formula,
                "tokens": [_formula_token_to_string(token) for token in tokens],
                "roundtrip_formula": printed,
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
            document = session["desktop"].loadComponentFromURL(
                file_url, "_blank", 0, load_props
            )
            if document is None:
                raise RuntimeError(f"LibreOffice could not open document: {ods_path}")
            return handler(request, session, document)
        finally:
            _close_document(document, save=False)


def _list_sheets(_request: dict[str, Any], _session: dict[str, Any], document: Any) -> dict[str, Any]:
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


def _execute_script(
    request: dict[str, Any], session: dict[str, Any], document: Any
) -> dict[str, Any]:
    script_uri = str(request["script_uri"])
    factory = session["component_context"].ServiceManager.createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory",
        session["component_context"],
    )
    script_provider = factory.createScriptProvider(document)
    script = script_provider.getScript(script_uri)
    return_value = script.invoke((), (), ())
    return {"executed": True, "return_value": _jsonable(return_value)}


def _click_form_button(
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

    execution_data = _execute_script({"script_uri": script_uri}, session, document)
    return {
        "clicked": True,
        "button_name": button_name,
        "script_uri": script_uri,
        "script_result": execution_data.get("return_value"),
    }


def _recalculate_document(
    _request: dict[str, Any], _session: dict[str, Any], document: Any
) -> dict[str, Any]:
    document.calculateAll()
    return {"recalculated": True}


class _OfficeSession:
    def __init__(self, request: dict[str, Any], use_gui: bool) -> None:
        self.request = request
        self.use_gui = use_gui
        self.process: subprocess.Popen[bytes] | None = None
        self.tmpdir: tempfile.TemporaryDirectory[str] | None = None
        self.data: dict[str, Any] = {}

    def __enter__(self) -> dict[str, Any]:
        import uno

        executable = str(self.request.get("office_executable") or "soffice")
        host = "127.0.0.1"
        port = _find_available_port()
        self.tmpdir = tempfile.TemporaryDirectory(prefix="xlsliberator-lo-worker-")
        profile_dir = Path(self.tmpdir.name) / "user"
        profile_dir.mkdir(parents=True, exist_ok=True)

        cmd = [
            executable,
            "--nologo",
            "--nodefault",
            "--norestore",
            "--nofirststartwizard",
            "--nolockcheck",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            f"--accept=socket,host={host},port={port};urp;",
        ]
        if not self.use_gui:
            cmd.insert(1, "--headless")

        env = dict(os.environ)
        self.process = subprocess.Popen(
            cmd,
            env=env,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        _wait_for_socket(host, port, int(self.request.get("start_timeout_seconds", 20)))

        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        component_context = resolver.resolve(
            f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
        )
        desktop = component_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", component_context
        )
        self.data = {
            "uno": uno,
            "component_context": component_context,
            "desktop": desktop,
            "host": host,
            "port": port,
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
            self.process = None
        if self.tmpdir is not None:
            self.tmpdir.cleanup()
            self.tmpdir = None


def _office_session(request: dict[str, Any], use_gui: bool) -> _OfficeSession:
    return _OfficeSession(request, use_gui)


def _property_value(name: str, value: Any) -> Any:
    from com.sun.star.beans import PropertyValue

    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def _set_parser_property(parser: Any, name: str, value: Any) -> None:
    with suppress(Exception):
        parser.setPropertyValue(name, value)


def _formula_token_to_string(token: Any) -> str:
    opcode = getattr(token, "OpCode", None)
    data = getattr(token, "Data", None)
    if data is None or data == "":
        return str(opcode)
    return f"{opcode}:{data}"


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


def _find_available_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _wait_for_socket(host: str, port: int, timeout_seconds: int) -> None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        if _process_exited_early(host, port):
            break
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.5)
            if sock.connect_ex((host, port)) == 0:
                return
        time.sleep(0.2)
    raise TimeoutError(f"LibreOffice UNO socket did not open on {host}:{port}")


def _process_exited_early(_host: str, _port: int) -> bool:
    return False


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
