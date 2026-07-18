"""Compatibility API delegating ODS script edits to the transactional package tool."""

from pathlib import Path
from typing import Any

from loguru import logger

from xlsliberator.validation_models import EventBindingIR


class MacroEmbedError(Exception):
    """Raised when macro embedding fails."""


def embed_python_macros(
    ods_path: str | Path,
    py_modules: dict[str, str],
    event_bindings: list[EventBindingIR] | None = None,
) -> list[EventBindingIR]:
    """Embed named Python modules through the transactional ODS package layer.

    Existing scripts, unknown package members, and unrelated manifest entries
    remain untouched. The operation either commits a verified package or leaves
    the original byte-for-byte unchanged.
    """
    if not py_modules:
        logger.warning("No Python modules to embed")
        return []

    from xlsliberator.odstool import OdsToolError, upsert_scripts

    logger.info(f"Embedding {len(py_modules)} Python modules into {ods_path}")
    try:
        result = upsert_scripts(
            ods_path,
            py_modules,
            event_bindings=event_bindings,
        )
    except OdsToolError as exc:
        raise MacroEmbedError(f"Failed to embed macros: {exc}") from exc
    logger.success(
        f"Embedded {len(py_modules)} Python modules transactionally ({result.after_sha256})"
    )
    return []


def remove_python_macros(ods_path: str | Path, module_names: list[str]) -> None:
    """Remove only explicitly named Python modules in one atomic transaction."""
    if not module_names:
        return

    from xlsliberator.odstool import OdsToolError, remove_scripts

    try:
        remove_scripts(ods_path, module_names)
    except OdsToolError as exc:
        raise MacroEmbedError(f"Failed to remove macros: {exc}") from exc


def attach_event_handler(
    doc: Any,
    event_name: str,
    script_url: str,
) -> None:
    """Attach an event handler to an open LibreOffice document."""
    from com.sun.star.beans import PropertyValue

    try:
        events = doc.getEvents()
        props = []

        prop_type = PropertyValue()
        prop_type.Name = "EventType"
        prop_type.Value = "Script"
        props.append(prop_type)

        prop_script = PropertyValue()
        prop_script.Name = "Script"
        prop_script.Value = script_url
        props.append(prop_script)

        from com.sun.star.uno import Any as UnoAny

        events.replaceByName(event_name, UnoAny("[]com.sun.star.beans.PropertyValue", tuple(props)))
        logger.debug(f"Attached event handler: {event_name} -> {script_url}")
    except Exception as exc:
        logger.warning(f"Failed to attach event handler {event_name}: {exc}")
        raise MacroEmbedError(f"Failed to attach event handler: {exc}") from exc


def create_on_open_marker_script() -> str:
    """Create the historical on-open marker script used by integration tests."""
    return '''# Auto-generated on_open event handler for testing
# Phase F6 - Gate G6 marker script

def on_open(*args):
    """Event handler called when document opens.

    Sets marker cell to verify event fired exactly once.
    """
    import uno

    # Get current document
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()

    # Get first sheet
    sheets = doc.getSheets()
    sheet = sheets.getByIndex(0)

    # Set marker cell A1
    cell = sheet.getCellByPosition(0, 0)  # A1
    cell.setString("OPEN_OK")

    return None
'''
