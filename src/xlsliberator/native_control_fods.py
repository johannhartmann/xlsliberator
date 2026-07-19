"""Build and augment ODF spreadsheets containing native LibreOffice controls.

The native form structure follows LibreOffice's own current Calc test fixture:

https://github.com/LibreOffice/core/blob/master/sc/qa/unit/tiledrendering/data/form-image-link.fods

Production workers first let the pinned LibreOffice runtime create a complete
ODS package, close it, and then use :func:`inject_native_buttons` to add only
the target-native form models and draw-page shapes.  The document is reopened
and persisted by LibreOffice before it is accepted.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

# Used only for element construction, namespace registration, and serialization.
# ODS input is parsed exclusively with defusedxml below.
import xml.etree.ElementTree as ElementTree  # nosec B405
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

from defusedxml.ElementTree import fromstring as safe_fromstring
from defusedxml.ElementTree import iterparse as safe_iterparse

_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
_NAMESPACES = {
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
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
    tag: str
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

    _register_document_namespaces(content)
    root = safe_fromstring(content)
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
                _qname("xlink", "href"): "",
                _qname("xlink", "type"): "simple",
            },
        )
        anchor_cell = table.find(".//table:table-cell", _NAMESPACES)
        if anchor_cell is None:
            row = ElementTree.SubElement(table, _qname("table", "table-row"))
            anchor_cell = ElementTree.SubElement(row, _qname("table", "table-cell"))

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
                    _qname("office", "string-value"): (
                        "com.sun.star.form.control.CommandButton"
                    ),
                },
            )
            ElementTree.SubElement(
                properties,
                _qname("form", "property"),
                {
                    _qname("form", "property-name"): "Tag",
                    _qname("office", "value-type"): "string",
                    _qname("office", "string-value"): button.tag,
                },
            )
            ElementTree.SubElement(
                anchor_cell,
                _qname("draw", "control"),
                {
                    _qname("draw", "control"): control_id,
                    _qname("draw", "name"): button.name,
                    _qname("draw", "z-index"): str(z_index),
                    _qname("svg", "x"): _cm(button.x),
                    _qname("svg", "y"): _cm(button.y),
                    _qname("svg", "width"): _cm(button.width),
                    _qname("svg", "height"): _cm(button.height),
                },
            )
        table.insert(0, forms)

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
    for _event, namespace in safe_iterparse(BytesIO(content), events=("start-ns",)):
        prefix, uri = namespace
        if not prefix.startswith("ns"):
            ElementTree.register_namespace(prefix, uri)
    for prefix, uri in _NAMESPACES.items():
        ElementTree.register_namespace(prefix, uri)


def _qname(prefix: str, local_name: str) -> str:
    return f"{{{_NAMESPACES[prefix]}}}{local_name}"


def _sheet_xml(
    sheet: NativeSheet,
    controls: list[tuple[NativeButton, str]],
) -> str:
    visibility = ' table:visibility="collapse"' if sheet.hidden else ""
    forms = _forms_xml(controls) if controls else ""
    shapes = "".join(_shape_xml(button, control_id) for button, control_id in controls)
    return f"""   <table:table table:name="{_xml_attr(sheet.name)}"{visibility}>
{forms}    <table:table-column table:number-columns-repeated="32"/>
    <table:table-row>
     <table:table-cell>
      <text:p/>
{shapes}\
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
      office:target-frame=""
      xlink:href=""
      xlink:type="simple">
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
    return f"""      <draw:control
       draw:z-index="0"
       draw:name="{_xml_attr(button.name)}"
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
