"""Build minimal ODF spreadsheets containing native LibreOffice controls.

The native form structure follows LibreOffice's own current Calc test fixture:

https://github.com/LibreOffice/core/blob/master/sc/qa/unit/tiledrendering/data/form-image-link.fods

The package writer avoids LibreOffice's unstable FODS-to-ODS ``storeAsURL``
path. The pinned Docker worker opens this valid ODS seed, populates it through
UNO, and persists it through the document's normal ``store`` lifecycle.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile

_MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"


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


def _sheet_xml(
    sheet: NativeSheet,
    controls: list[tuple[NativeButton, str]],
) -> str:
    visibility = ' table:visibility="collapse"' if sheet.hidden else ""
    forms = _forms_xml(controls) if controls else ""
    shapes = "".join(_shape_xml(button, control_id) for button, control_id in controls)
    shapes_xml = f"    <table:shapes>{shapes}    </table:shapes>\n" if shapes else ""
    return f"""   <table:table table:name="{_xml_attr(sheet.name)}"{visibility}>
{forms}    <table:table-column table:number-columns-repeated="32"/>
    <table:table-row>
     <table:table-cell>
      <text:p/>
     </table:table-cell>
    </table:table-row>
{shapes_xml}\
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
       xlink:href=""
       form:image-data=""
       form:delay-for-repeat="PT0.050000000S"
       form:image-position="center"
       office:target-frame="">
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


__all__ = ["NativeButton", "NativeSheet", "write_native_button_seed"]
