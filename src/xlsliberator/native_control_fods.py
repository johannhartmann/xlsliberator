"""Build and augment ODF spreadsheets containing native LibreOffice controls.

The form model and export-safe draw-page placement follow LibreOffice's current
Calc fixtures and round-trip export test:

https://github.com/LibreOffice/core/blob/master/sc/qa/unit/tiledrendering/data/form-image-link.fods
https://github.com/LibreOffice/core/blob/master/sc/qa/unit/data/fods/shapes_foreground_background.fods

Production workers first let the pinned LibreOffice runtime create a complete
ODS package, close it, and then use :func:`inject_native_buttons` to add only
the target-native form models and draw-page shapes.  The document is reopened
and persisted by LibreOffice before it is accepted.
"""

from __future__ import annotations

# This module runs in LibreOffice's bundled, standard-library-only Python.
# XML declarations which make ElementTree unsafe are rejected before parsing.
import xml.etree.ElementTree as ElementTree  # nosec B405
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
_MAX_CONTENT_XML_BYTES = 32 * 1024 * 1024
_NAMESPACES = {
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "ooo": "http://openoffice.org/2004/office",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "xml": "http://www.w3.org/XML/1998/namespace",
}


@dataclass(frozen=True)
class NativeButton:
    """A native command-button model and its Calc draw-page placement."""

    name: str
    label: str
    x: int
    y: int
    width: int
    height: int = 1_200


@dataclass(frozen=True)
class NativeSheet:
    """A sheet in the intermediate ODS document."""

    name: str
    buttons: tuple[NativeButton, ...] = ()
    hidden: bool = False


def _xml_attr(value: str) -> str:
    """Encode text for a double-quoted XML attribute without parsing XML."""
    return (
        value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;").replace(">", "&gt;")
    )


def write_native_button_seed(path: Path, sheets: tuple[NativeSheet, ...]) -> None:
    """Write a minimal ODS spreadsheet with sheet-local native buttons."""
    if not sheets:
        raise ValueError("native-control seed requires at least one sheet")
    if all(sheet.hidden for sheet in sheets):
        raise ValueError("native-control seed requires a visible sheet")

    control_number = 0
    sheet_xml: list[str] = []
    for sheet in sheets:
        controls: list[tuple[NativeButton, str]] = []
        for button in sheet.buttons:
            control_number += 1
            controls.append((button, f"control{control_number}"))
        sheet_xml.append(_sheet_xml(sheet, controls))

    content = f"""<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 xmlns:ooo="http://openoffice.org/2004/office"
 office:version="1.4">
 <office:automatic-styles>
  <style:style style:name="XLSLiberatorControlShape" style:family="graphic">
   <style:graphic-properties fo:border="none"/>
  </style:style>
  <style:style style:name="XLSLiberatorControlText" style:family="paragraph">
   <style:text-properties fo:font-size="10pt"/>
  </style:style>
 </office:automatic-styles>
 <office:body>
  <office:spreadsheet>
{"".join(sheet_xml)}
  </office:spreadsheet>
 </office:body>
</office:document-content>
"""
    styles = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 office:version="1.4">
 <office:styles/>
</office:document-styles>
"""
    manifest = f"""<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
 xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
 manifest:version="1.4">
 <manifest:file-entry
  manifest:full-path="/"
  manifest:version="1.4"
  manifest:media-type="{_MIMETYPE}"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
</manifest:manifest>
"""
    with ZipFile(path, "w") as package:
        package.writestr("mimetype", _MIMETYPE, compress_type=ZIP_STORED)
        package.writestr("META-INF/manifest.xml", manifest, compress_type=ZIP_DEFLATED)
        package.writestr("content.xml", content, compress_type=ZIP_DEFLATED)
        package.writestr("styles.xml", styles, compress_type=ZIP_DEFLATED)


def inject_native_buttons(path: Path, sheets: tuple[NativeSheet, ...]) -> None:
    """Inject native button models into a closed LibreOffice-created ODS.

    All existing package members and their ZIP metadata are retained.  Only
    ``content.xml`` is rewritten, and callers must round-trip the result
    through the pinned LibreOffice runtime before treating it as usable.
    """
    if not path.is_file():
        raise FileNotFoundError(f"native-control ODS does not exist: {path}")
    requested = tuple(sheet for sheet in sheets if sheet.buttons)
    if not requested:
        raise ValueError("native-control injection requires at least one button")

    with ZipFile(path) as package:
        infos = package.infolist()
        members = {info.filename: package.read(info.filename) for info in infos}
    if members.get("mimetype") != _MIMETYPE.encode():
        raise ValueError("native controls can be injected only into an ODS package")
    content = members.get("content.xml")
    if content is None:
        raise ValueError("ODS package has no content.xml")

    _validate_content_xml(content)
    _register_document_namespaces(content)
    # Safe because size, DTD, and entity declarations were rejected above.
    root = ElementTree.fromstring(content)  # nosec B314
    _ensure_control_styles(root)
    table_name = _qname("table", "name")
    tables = {
        str(table.attrib.get(table_name, "")): table
        for table in root.findall(".//table:table", _NAMESPACES)
    }
    existing_ids = {
        value
        for element in root.iter()
        for name, value in element.attrib.items()
        if name in {_qname("form", "id"), _qname("xml", "id")}
    }
    next_control = 1
    for native_sheet in requested:
        table = tables.get(native_sheet.name)
        if table is None:
            raise ValueError(f"native-control sheet is missing from ODS: {native_sheet.name}")
        if table.find("./office:forms", _NAMESPACES) is not None:
            raise ValueError(f"native-control sheet already contains forms: {native_sheet.name}")

        forms = ElementTree.Element(
            _qname("office", "forms"),
            {
                _qname("form", "automatic-focus"): "false",
                _qname("form", "apply-design-mode"): "false",
            },
        )
        form = ElementTree.SubElement(
            forms,
            _qname("form", "form"),
            {
                _qname("form", "name"): f"XLSLiberatorForm{next_control}",
                _qname("form", "apply-filter"): "true",
                _qname("form", "command-type"): "table",
                _qname("form", "control-implementation"): ("ooo:com.sun.star.form.component.Form"),
                _qname("office", "target-frame"): "",
            },
        )
        form_properties = ElementTree.SubElement(form, _qname("form", "properties"))
        ElementTree.SubElement(
            form_properties,
            _qname("form", "property"),
            {
                _qname("form", "property-name"): "PropertyChangeNotificationEnabled",
                _qname("office", "value-type"): "boolean",
                _qname("office", "boolean-value"): "true",
            },
        )
        ElementTree.SubElement(
            form_properties,
            _qname("form", "property"),
            {
                _qname("form", "property-name"): "TargetURL",
                _qname("office", "value-type"): "string",
                _qname("office", "string-value"): "",
            },
        )
        shapes = table.find("./table:shapes", _NAMESPACES)
        created_shapes = shapes is None
        if shapes is None:
            shapes = ElementTree.Element(_qname("table", "shapes"))

        for z_index, button in enumerate(native_sheet.buttons):
            while f"control{next_control}" in existing_ids:
                next_control += 1
            control_id = f"control{next_control}"
            existing_ids.add(control_id)
            next_control += 1
            model = ElementTree.SubElement(
                form,
                _qname("form", "button"),
                {
                    _qname("form", "id"): control_id,
                    _qname("xml", "id"): control_id,
                    _qname("form", "name"): button.name,
                    _qname("form", "control-implementation"): (
                        "ooo:com.sun.star.form.component.CommandButton"
                    ),
                    _qname("form", "label"): button.label,
                    _qname("office", "target-frame"): "",
                    _qname("xlink", "href"): "",
                    _qname("form", "image-data"): "",
                    _qname("form", "delay-for-repeat"): "PT0.050000000S",
                    _qname("form", "image-position"): "center",
                },
            )
            properties = ElementTree.SubElement(model, _qname("form", "properties"))
            ElementTree.SubElement(
                properties,
                _qname("form", "property"),
                {
                    _qname("form", "property-name"): "DefaultControl",
                    _qname("office", "value-type"): "string",
                    _qname("office", "string-value"): ("com.sun.star.form.control.CommandButton"),
                },
            )
            ElementTree.SubElement(
                shapes,
                _qname("draw", "control"),
                {
                    _qname("draw", "control"): control_id,
                    _qname("draw", "name"): button.name,
                    _qname("draw", "style-name"): "XLSLiberatorControlShape",
                    _qname("draw", "text-style-name"): "XLSLiberatorControlText",
                    _qname("draw", "z-index"): str(z_index),
                    _qname("svg", "x"): _cm(button.x),
                    _qname("svg", "y"): _cm(button.y),
                    _qname("svg", "width"): _cm(button.width),
                    _qname("svg", "height"): _cm(button.height),
                },
            )
        table.insert(0, forms)
        if created_shapes:
            table.insert(1, shapes)

    # ElementTree otherwise drops this declaration when ``ooo`` is used only in
    # QName-valued attributes.  LibreOffice needs it to resolve form services.
    if not _tree_uses_namespace(root, _NAMESPACES["ooo"]):
        root.set("xmlns:ooo", _NAMESPACES["ooo"])
    members["content.xml"] = ElementTree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )
    temporary = path.with_name(f".{path.name}.native-controls")
    temporary.unlink(missing_ok=True)
    try:
        with ZipFile(temporary, "w") as package:
            for info in infos:
                package.writestr(info, members[info.filename])
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def _register_document_namespaces(content: bytes) -> None:
    # Safe because inject_native_buttons validates content before calling this helper.
    for _event, namespace in ElementTree.iterparse(  # nosec B314
        BytesIO(content),
        events=("start-ns",),
    ):
        prefix, uri = namespace
        if not prefix.startswith("ns"):
            ElementTree.register_namespace(prefix, uri)
    for prefix, uri in _NAMESPACES.items():
        ElementTree.register_namespace(prefix, uri)


def _validate_content_xml(content: bytes) -> None:
    if len(content) > _MAX_CONTENT_XML_BYTES:
        raise ValueError(f"ODS content.xml exceeds the {_MAX_CONTENT_XML_BYTES}-byte safety limit")
    if b"\x00" in content:
        raise ValueError("ODS content.xml must use an ASCII-compatible XML encoding")
    normalized = content.upper()
    if b"<!DOCTYPE" in normalized or b"<!ENTITY" in normalized:
        raise ValueError("ODS content.xml must not contain DTD or entity declarations")


def _tree_uses_namespace(root: ElementTree.Element, namespace: str) -> bool:
    expanded_prefix = f"{{{namespace}}}"
    return any(
        (isinstance(element.tag, str) and element.tag.startswith(expanded_prefix))
        or any(name.startswith(expanded_prefix) for name in element.attrib)
        for element in root.iter()
    )


def _ensure_control_styles(root: ElementTree.Element) -> None:
    automatic_styles = root.find("./office:automatic-styles", _NAMESPACES)
    if automatic_styles is None:
        automatic_styles = ElementTree.Element(_qname("office", "automatic-styles"))
        body = root.find("./office:body", _NAMESPACES)
        root.insert(list(root).index(body) if body is not None else 0, automatic_styles)

    style_name = _qname("style", "name")
    existing_names = {
        style.attrib.get(style_name)
        for style in automatic_styles.findall("./style:style", _NAMESPACES)
    }
    if "XLSLiberatorControlShape" not in existing_names:
        shape_style = ElementTree.SubElement(
            automatic_styles,
            _qname("style", "style"),
            {
                style_name: "XLSLiberatorControlShape",
                _qname("style", "family"): "graphic",
            },
        )
        ElementTree.SubElement(
            shape_style,
            _qname("style", "graphic-properties"),
            {_qname("fo", "border"): "none"},
        )
    if "XLSLiberatorControlText" not in existing_names:
        text_style = ElementTree.SubElement(
            automatic_styles,
            _qname("style", "style"),
            {
                style_name: "XLSLiberatorControlText",
                _qname("style", "family"): "paragraph",
            },
        )
        ElementTree.SubElement(
            text_style,
            _qname("style", "text-properties"),
            {_qname("fo", "font-size"): "10pt"},
        )


def _qname(prefix: str, local_name: str) -> str:
    return f"{{{_NAMESPACES[prefix]}}}{local_name}"


def _sheet_xml(
    sheet: NativeSheet,
    controls: list[tuple[NativeButton, str]],
) -> str:
    visibility = ' table:visibility="collapse"' if sheet.hidden else ""
    forms = _forms_xml(controls) if controls else ""
    shapes = (
        "    <table:shapes>\n"
        + "".join(_shape_xml(button, control_id) for button, control_id in controls)
        + "    </table:shapes>\n"
        if controls
        else ""
    )
    return f"""   <table:table table:name="{_xml_attr(sheet.name)}"{visibility}>
{forms}{shapes}    <table:table-column table:number-columns-repeated="32"/>
    <table:table-row>
     <table:table-cell>
      <text:p/>
     </table:table-cell>
    </table:table-row>
   </table:table>
"""


def _forms_xml(controls: list[tuple[NativeButton, str]]) -> str:
    buttons = "".join(_button_xml(button, control_id) for button, control_id in controls)
    return f"""    <office:forms form:automatic-focus="false" form:apply-design-mode="false">
     <form:form
      form:name="CertificationForm"
      form:apply-filter="true"
      form:command-type="table"
      form:control-implementation="ooo:com.sun.star.form.component.Form"
      office:target-frame="">
      <form:properties>
       <form:property
        form:property-name="PropertyChangeNotificationEnabled"
        office:value-type="boolean"
        office:boolean-value="true"/>
       <form:property
        form:property-name="TargetURL"
        office:value-type="string"
        office:string-value=""/>
      </form:properties>
{buttons}     </form:form>
    </office:forms>
"""


def _button_xml(button: NativeButton, control_id: str) -> str:
    name = _xml_attr(button.name)
    label = _xml_attr(button.label)
    return f"""      <form:button
       form:name="{name}"
       form:control-implementation="ooo:com.sun.star.form.component.CommandButton"
       xml:id="{control_id}"
       form:id="{control_id}"
       form:label="{label}"
       office:target-frame=""
       xlink:href=""
       form:image-data=""
       form:delay-for-repeat="PT0.050000000S"
       form:image-position="center">
       <form:properties>
        <form:property
         form:property-name="DefaultControl"
         office:value-type="string"
         office:string-value="com.sun.star.form.control.CommandButton"/>
       </form:properties>
      </form:button>
"""


def _shape_xml(button: NativeButton, control_id: str) -> str:
    return f"""     <draw:control
       draw:z-index="0"
       draw:name="{_xml_attr(button.name)}"
       draw:style-name="XLSLiberatorControlShape"
       draw:text-style-name="XLSLiberatorControlText"
       svg:width="{_cm(button.width)}"
       svg:height="{_cm(button.height)}"
       svg:x="{_cm(button.x)}"
       svg:y="{_cm(button.y)}"
       draw:control="{control_id}"/>
"""


def _cm(value: int) -> str:
    if value < 0:
        raise ValueError("native-control geometry must not be negative")
    return f"{value / 1000:g}cm"


__all__ = [
    "NativeButton",
    "NativeSheet",
    "inject_native_buttons",
    "write_native_button_seed",
]
