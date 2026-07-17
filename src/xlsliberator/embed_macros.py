"""Transactional Python macro upserts for ODS packages."""

import ast
import os
import tempfile

# Used only for construction/serialization; all untrusted parsing uses defusedxml.
import xml.etree.ElementTree as ET  # nosec B405
import zipfile
from pathlib import Path
from typing import Any

from defusedxml.ElementTree import fromstring as safe_fromstring
from loguru import logger

from xlsliberator.validation_models import EventBindingIR


class MacroEmbedError(Exception):
    """Raised when macro embedding fails."""


def embed_python_macros(
    ods_path: str | Path,
    py_modules: dict[str, str],
    event_bindings: list[EventBindingIR] | None = None,
) -> list[EventBindingIR]:
    """Embed Python macros into an ODS file.

    Args:
        ods_path: Path to ODS file
        py_modules: Dict mapping module names to Python source code
                   e.g., {"doc_events.py": "def on_open(doc): ..."}
        event_bindings: Optional source-map-aware event bindings. When provided
                   they drive the rewrite (taking precedence over heuristic
                   module matching); otherwise the legacy heuristic is used.

    Returns:
        The event bindings that could not be rewritten (empty when all resolved
        or when the heuristic path is used).

    Raises:
        MacroEmbedError: If embedding fails

    Only the named modules are replaced. Existing scripts, unknown package
    members, and manifest entries remain untouched.
    """
    ods_path = Path(ods_path)

    if not ods_path.exists():
        raise MacroEmbedError(f"ODS file not found: {ods_path}")

    if not py_modules:
        logger.warning("No Python modules to embed")
        return []

    logger.info(f"Embedding {len(py_modules)} Python modules into {ods_path}")

    normalized_modules = _normalize_modules(py_modules)
    unresolved: list[EventBindingIR] = []
    temp_path: Path | None = None
    try:
        descriptor, raw_temp_path = tempfile.mkstemp(
            prefix=f".{ods_path.name}.", suffix=".tmp", dir=ods_path.parent
        )
        os.close(descriptor)
        temp_path = Path(raw_temp_path)

        with (
            zipfile.ZipFile(ods_path, "r") as zip_in,
            zipfile.ZipFile(temp_path, "w") as zip_out,
        ):
            owned_paths = {f"Scripts/python/{name}" for name in normalized_modules}
            for item in zip_in.infolist():
                if item.filename == "META-INF/manifest.xml":
                    continue
                if item.filename in owned_paths:
                    continue
                if item.filename == "content.xml":
                    from xlsliberator.event_binding_writer import rewrite_event_bindings

                    data = zip_in.read(item.filename)
                    updated_content, unresolved = rewrite_event_bindings(
                        data.decode("utf-8"), normalized_modules, event_bindings
                    )
                    if unresolved:
                        raise MacroEmbedError(
                            f"{len(unresolved)} event binding(s) could not be rewritten"
                        )
                    zip_out.writestr(item, updated_content.encode("utf-8"))
                    continue
                zip_out.writestr(item, zip_in.read(item.filename))

            for module_name, module_code in normalized_modules.items():
                script_path = f"Scripts/python/{module_name}"
                zip_out.writestr(script_path, module_code.encode("utf-8"))
                logger.debug(f"Added Python module: {script_path}")

            manifest_content = _create_manifest_with_scripts(
                zip_in, list(normalized_modules), remove_modules=[]
            )
            zip_out.writestr("META-INF/manifest.xml", manifest_content)

        _fsync_file(temp_path)
        _validate_embedded_package(temp_path, normalized_modules, event_bindings)
        os.replace(temp_path, ods_path)
        _fsync_directory(ods_path.parent)
        temp_path = None
        logger.success(f"Embedded {len(normalized_modules)} Python modules successfully")

    except Exception as exc:
        if temp_path is not None:
            temp_path.unlink(missing_ok=True)
        if isinstance(exc, MacroEmbedError):
            raise
        raise MacroEmbedError(f"Failed to embed macros: {exc}") from exc

    return unresolved


def remove_python_macros(ods_path: str | Path, module_names: list[str]) -> None:
    """Remove only explicitly named Python modules in one atomic transaction."""
    path = Path(ods_path)
    removal = _normalize_module_names(module_names)
    if not removal:
        return
    descriptor, raw_temp_path = tempfile.mkstemp(
        prefix=f".{path.name}.", suffix=".tmp", dir=path.parent
    )
    os.close(descriptor)
    temp_path = Path(raw_temp_path)
    try:
        removed_paths = {f"Scripts/python/{name}" for name in removal}
        with zipfile.ZipFile(path, "r") as zip_in, zipfile.ZipFile(temp_path, "w") as zip_out:
            for item in zip_in.infolist():
                if item.filename == "META-INF/manifest.xml" or item.filename in removed_paths:
                    continue
                zip_out.writestr(item, zip_in.read(item.filename))
            manifest = _create_manifest_with_scripts(
                zip_in, module_names=[], remove_modules=removal
            )
            zip_out.writestr("META-INF/manifest.xml", manifest)
        _fsync_file(temp_path)
        _validate_embedded_package(temp_path, {}, None)
        os.replace(temp_path, path)
        _fsync_directory(path.parent)
    except Exception as exc:
        temp_path.unlink(missing_ok=True)
        if isinstance(exc, MacroEmbedError):
            raise
        raise MacroEmbedError(f"Failed to remove macros: {exc}") from exc


def _normalize_modules(py_modules: dict[str, str]) -> dict[str, str]:
    names = _normalize_module_names(list(py_modules))
    normalized: dict[str, str] = {}
    for original, name in zip(py_modules, names, strict=True):
        source = py_modules[original]
        try:
            compile(source, name, "exec")
        except (SyntaxError, ValueError) as exc:
            raise MacroEmbedError(f"Invalid Python module {name}: {exc}") from exc
        normalized[name] = source
    return normalized


def _normalize_module_names(module_names: list[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for raw_name in module_names:
        name = raw_name if raw_name.endswith(".py") else f"{raw_name}.py"
        if not name or Path(name).name != name or "/" in name or "\\" in name:
            raise MacroEmbedError(f"Invalid Python module name: {raw_name!r}")
        collision_key = name.casefold()
        if collision_key in seen:
            raise MacroEmbedError(f"Duplicate Python module name: {name}")
        seen.add(collision_key)
        normalized.append(name)
    return normalized


def _create_manifest_with_scripts(
    zip_in: zipfile.ZipFile,
    module_names: list[str],
    remove_modules: list[str],
) -> str:
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
        root = safe_fromstring(manifest_data)
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

    owned_paths = {
        *(f"Scripts/python/{name}" for name in module_names),
        *(f"Scripts/python/{name}" for name in remove_modules),
    }
    for entry in root.findall(f".//{{{NS['manifest']}}}file-entry"):
        path = entry.get(f"{{{NS['manifest']}}}full-path", "")
        if path in owned_paths:
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
    )

    return buffer.getvalue().decode("UTF-8")


def _validate_embedded_package(
    path: Path,
    expected_modules: dict[str, str],
    event_bindings: list[EventBindingIR] | None,
) -> None:
    """Fail before replacement unless package, scripts, manifest, and bindings agree."""
    try:
        with zipfile.ZipFile(path, "r") as archive:
            if archive.testzip() is not None:
                raise MacroEmbedError("ODS ZIP contains a corrupt member")
            infos = archive.infolist()
            if not infos or infos[0].filename != "mimetype":
                raise MacroEmbedError("ODS mimetype must be the first package member")
            if infos[0].compress_type != zipfile.ZIP_STORED:
                raise MacroEmbedError("ODS mimetype must be stored without compression")
            expected_mimetype = b"application/vnd.oasis.opendocument.spreadsheet"
            if archive.read("mimetype") != expected_mimetype:
                raise MacroEmbedError("ODS mimetype is invalid")
            names = set(archive.namelist())
            for required in ("content.xml", "META-INF/manifest.xml"):
                if required not in names:
                    raise MacroEmbedError(f"ODS package is missing {required}")

            manifest_root = safe_fromstring(archive.read("META-INF/manifest.xml"))
            namespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
            manifest_paths = {
                entry.get(f"{{{namespace}}}full-path", "")
                for entry in manifest_root.findall(f".//{{{namespace}}}file-entry")
            }
            scripts: dict[str, set[str]] = {}
            for member in sorted(name for name in names if name.startswith("Scripts/python/")):
                if member.endswith("/"):
                    continue
                if member not in manifest_paths:
                    raise MacroEmbedError(f"Manifest omits embedded script {member}")
                source = archive.read(member).decode("utf-8")
                tree = ast.parse(source, filename=member)
                scripts[Path(member).name] = {
                    node.name
                    for node in tree.body
                    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
                }
            dangling_scripts = {
                name
                for name in manifest_paths
                if name.startswith("Scripts/python/")
                and not name.endswith("/")
                and name not in names
            }
            if dangling_scripts:
                raise MacroEmbedError(
                    f"Manifest references missing scripts: {sorted(dangling_scripts)}"
                )
            for module_name, source in expected_modules.items():
                member = f"Scripts/python/{module_name}"
                if archive.read(member).decode("utf-8") != source:
                    raise MacroEmbedError(
                        f"Embedded module differs from requested source: {module_name}"
                    )
            _validate_event_targets(event_bindings, scripts)
    except (KeyError, UnicodeDecodeError, zipfile.BadZipFile, ET.ParseError) as exc:
        raise MacroEmbedError(f"Invalid ODS package: {exc}") from exc


def _validate_event_targets(
    event_bindings: list[EventBindingIR] | None,
    scripts: dict[str, set[str]],
) -> None:
    for binding in event_bindings or []:
        target = binding.target_script_uri
        if not target:
            raise MacroEmbedError(f"Event binding {binding.id} has no target")
        marker = "vnd.sun.star.script:"
        if not target.startswith(marker) or "$" not in target:
            raise MacroEmbedError(f"Event binding {binding.id} has an invalid target URI")
        module, procedure = target[len(marker) :].split("$", 1)
        procedure = procedure.split("?", 1)[0]
        if module not in scripts or procedure not in scripts[module]:
            raise MacroEmbedError(
                f"Event binding {binding.id} targets missing {module}${procedure}"
            )


def _fsync_file(path: Path) -> None:
    with path.open("rb") as handle:
        os.fsync(handle.fileno())


def _fsync_directory(path: Path) -> None:
    descriptor = os.open(path, os.O_RDONLY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


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
