"""Python macro embedding into ODS files (Phase F6)."""

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any

from loguru import logger


class MacroEmbedError(Exception):
    """Raised when macro embedding fails."""


def embed_python_macros(
    ods_path: str | Path,
    py_modules: dict[str, str],
) -> None:
    """Embed Python macros into an ODS file.

    Args:
        ods_path: Path to ODS file
        py_modules: Dict mapping module names to Python source code
                   e.g., {"doc_events.py": "def on_open(doc): ..."}

    Raises:
        MacroEmbedError: If embedding fails

    Note:
        Phase F6 implementation - embeds Python scripts into Scripts/python/
        and updates META-INF/manifest.xml. Also rewrites VBA event handlers
        to point to Python-UNO equivalents.
    """
    ods_path = Path(ods_path)

    if not ods_path.exists():
        raise MacroEmbedError(f"ODS file not found: {ods_path}")

    if not py_modules:
        logger.warning("No Python modules to embed")
        return

    logger.info(f"Embedding {len(py_modules)} Python modules into {ods_path}")

    try:
        # Create temporary copy to work with
        temp_path = ods_path.with_suffix(".ods.tmp")

        # Open original ODS as ZIP and create new ZIP with embedded macros
        with (
            zipfile.ZipFile(ods_path, "r") as zip_in,
            zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zip_out,
        ):
            # Copy existing files
            for item in zip_in.infolist():
                # Skip META-INF/manifest.xml - we'll recreate it
                if item.filename == "META-INF/manifest.xml":
                    continue
                # Skip any existing Scripts/python/ files
                if item.filename.startswith("Scripts/python/"):
                    continue
                # Process content.xml to rewrite VBA event handlers
                if item.filename == "content.xml":
                    data = zip_in.read(item.filename)
                    updated_content = _rewrite_vba_to_python_event_handlers(
                        data.decode("utf-8"), py_modules
                    )
                    zip_out.writestr(item, updated_content.encode("utf-8"))
                    continue

                data = zip_in.read(item.filename)
                zip_out.writestr(item, data)

            # Add Python modules
            for module_name, module_code in py_modules.items():
                script_path = f"Scripts/python/{module_name}"
                zip_out.writestr(script_path, module_code)
                logger.debug(f"Added Python module: {script_path}")

            # Update manifest.xml
            manifest_content = _create_manifest_with_scripts(zip_in, list(py_modules.keys()))
            zip_out.writestr("META-INF/manifest.xml", manifest_content)

        # Replace original with updated file
        temp_path.replace(ods_path)
        logger.success(f"Embedded {len(py_modules)} Python modules successfully")

    except Exception as e:
        raise MacroEmbedError(f"Failed to embed macros: {e}") from e


def _rewrite_vba_to_python_event_handlers(content_xml: str, py_modules: dict[str, str]) -> str:
    """Rewrite VBA event handlers to point to Python-UNO equivalents.

    Args:
        content_xml: Content of content.xml as string
        py_modules: Dict mapping Python module names to source code

    Returns:
        Updated content.xml with Python event handlers

    Note:
        Converts VBA event handlers like:
        vnd.sun.star.script:VBAProject.Game.StartButton_Click?language=Basic&location=document
        To Python equivalents:
        vnd.sun.star.script:Game.bas.py$StartButton_Click?language=Python&location=document
    """
    import re

    # Build mapping from VBA module names to Python module names
    # e.g., "Game.bas" -> "Game.bas.py" AND "Game" -> "Game.bas.py"
    vba_to_py: dict[str, str] = {}
    for py_module_name in py_modules:
        # Remove .py extension to get base name (e.g., "Game.bas.py" -> "Game.bas")
        if py_module_name.endswith(".py"):
            base_name = py_module_name[:-3]
            vba_to_py[base_name] = py_module_name

            # Also map without VBA extension (e.g., "Game.bas" -> "Game")
            # VBAProject.Game.Function should map to Game.bas.py
            if "." in base_name:
                module_only = base_name.split(".")[0]
                vba_to_py[module_only] = py_module_name

    if not vba_to_py:
        logger.debug("No VBA modules to map, skipping event handler rewriting")
        return content_xml

    # Pattern to match VBA script URLs in event handlers
    # Example: vnd.sun.star.script:VBAProject.Game.StartButton_Click?language=Basic&location=document
    pattern = (
        r"(vnd\.sun\.star\.script:)VBAProject\.([^?]+)\?language=Basic(&amp;location=document)"
    )

    def replace_handler(match: re.Match[str]) -> str:
        """Replace VBA handler with Python equivalent."""
        prefix: str = match.group(1)  # "vnd.sun.star.script:"
        vba_path: str = match.group(2)  # "Game.StartButton_Click"
        location: str = match.group(3)  # "&amp;location=document"
        original: str = match.group(0)  # Full match

        # Split VBA path into module and function
        # "Game.StartButton_Click" -> ["Game", "StartButton_Click"]
        parts = vba_path.split(".")
        if len(parts) < 2:
            logger.warning(f"Invalid VBA path format: {vba_path}")
            return original  # Return original if malformed

        # Find matching Python module
        # Try various combinations: "Game", "Game.bas", etc.
        py_module = None
        function_name = ""
        for i in range(len(parts) - 1, 0, -1):
            vba_module = ".".join(parts[:i])
            if vba_module in vba_to_py:
                py_module = vba_to_py[vba_module]
                function_name = ".".join(parts[i:])
                break

        if not py_module:
            logger.warning(f"No Python module found for VBA path: {vba_path}")
            return original  # Keep VBA handler if no Python equivalent

        # Build Python script URL
        # Format: Module.py$function_name
        python_url = f"{prefix}{py_module}${function_name}?language=Python{location}"
        logger.debug(f"Rewrote event handler: {original} -> {python_url}")
        return python_url

    # Replace all VBA handlers with Python equivalents
    updated_content = re.sub(pattern, replace_handler, content_xml)

    # Count replacements
    replacements = len(re.findall(pattern, content_xml))
    if replacements > 0:
        logger.info(f"Rewrote {replacements} VBA event handlers to Python")

    return updated_content


def _create_manifest_with_scripts(zip_in: zipfile.ZipFile, module_names: list[str]) -> str:
    """Create updated manifest.xml with Python script entries.

    Args:
        zip_in: Input ZIP file (original ODS)
        module_names: List of Python module names to add

    Returns:
        Updated manifest.xml content as string
    """
    # Define XML namespaces
    NS = {
        "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    }

    # Register namespace
    ET.register_namespace("manifest", NS["manifest"])

    # Try to read existing manifest, or create new one
    try:
        manifest_data = zip_in.read("META-INF/manifest.xml")
        # Safe: parsing manifest from ODS file we control
        root = ET.fromstring(manifest_data)  # nosec B314
    except KeyError:
        # No existing manifest - create new one
        root = ET.Element(f"{{{NS['manifest']}}}manifest")
        root.set(f"{{{NS['manifest']}}}version", "1.3")

        # Add standard ODS entries
        _add_manifest_entry(root, "/", "application/vnd.oasis.opendocument.spreadsheet")
        _add_manifest_entry(root, "content.xml", "text/xml")
        _add_manifest_entry(root, "styles.xml", "text/xml")
        _add_manifest_entry(root, "meta.xml", "text/xml")
        _add_manifest_entry(root, "settings.xml", "text/xml")

    # Remove any existing Scripts/python/ entries
    for entry in root.findall(f".//{{{NS['manifest']}}}file-entry"):
        path = entry.get(f"{{{NS['manifest']}}}full-path", "")
        if path.startswith("Scripts/python/"):
            root.remove(entry)

    # Add Scripts/python/ directory entry
    scripts_dir_exists = False
    for entry in root.findall(f".//{{{NS['manifest']}}}file-entry"):
        path = entry.get(f"{{{NS['manifest']}}}full-path", "")
        if path == "Scripts/python/":
            scripts_dir_exists = True
            break

    if not scripts_dir_exists:
        _add_manifest_entry(root, "Scripts/python/", "application/binary")

    # Add Python script entries
    for module_name in module_names:
        script_path = f"Scripts/python/{module_name}"
        _add_manifest_entry(root, script_path, "application/binary")

    # Convert to string with XML declaration
    tree = ET.ElementTree(root)
    import io

    buffer = io.BytesIO()
    tree.write(
        buffer,
        encoding="UTF-8",
        xml_declaration=True,
        default_namespace=NS["manifest"],
    )

    return buffer.getvalue().decode("UTF-8")


def _add_manifest_entry(root: ET.Element, full_path: str, media_type: str) -> None:
    """Add a file entry to manifest.

    Args:
        root: Manifest root element
        full_path: Full path of file in ODS
        media_type: MIME type of file
    """
    NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"

    entry = ET.SubElement(root, f"{{{NS}}}file-entry")
    entry.set(f"{{{NS}}}full-path", full_path)
    entry.set(f"{{{NS}}}media-type", media_type)


def attach_event_handler(
    doc: Any,
    event_name: str,
    script_url: str,
) -> None:
    """Attach event handler to LibreOffice document.

    Args:
        doc: LibreOffice document
        event_name: Event name (e.g., "OnLoad")
        script_url: Script URL (e.g., "vnd.sun.star.script:doc_events.py$on_open?...")

    Note:
        This configures the document to call the Python script when event fires.
        Phase F6 implementation focuses on OnLoad event.
    """
    from com.sun.star.beans import PropertyValue

    try:
        # Get document events
        events = doc.getEvents()

        # Create event binding
        props = []

        # EventType: "Script"
        prop_type = PropertyValue()
        prop_type.Name = "EventType"
        prop_type.Value = "Script"
        props.append(prop_type)

        # Script: URL to Python function
        prop_script = PropertyValue()
        prop_script.Name = "Script"
        prop_script.Value = script_url
        props.append(prop_script)

        # Attach event
        from com.sun.star.uno import Any as UnoAny

        events.replaceByName(event_name, UnoAny("[]com.sun.star.beans.PropertyValue", tuple(props)))

        logger.debug(f"Attached event handler: {event_name} -> {script_url}")

    except Exception as e:
        logger.warning(f"Failed to attach event handler {event_name}: {e}")
        raise MacroEmbedError(f"Failed to attach event handler: {e}") from e


def create_on_open_marker_script() -> str:
    """Create a simple on_open script that sets a marker cell.

    Returns:
        Python source code for on_open event handler

    Note:
        This is a test script for Phase F6 Gate G6 verification.
        Sets cell Sheet1.A1 to "OPEN_OK" when document opens.
    """
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
