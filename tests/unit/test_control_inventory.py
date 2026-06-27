"""Tests for deterministic ODS control inventory."""

import zipfile
from pathlib import Path

from xlsliberator.control_inventory import (
    extract_controls_from_ods,
    extract_event_bindings_from_ods,
)


def _write_ods(path: Path, content_xml: str) -> None:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("content.xml", content_xml)


def test_extract_controls_and_event_bindings_from_ods(tmp_path: Path) -> None:
    """ODS XML parsing should return controls and event bindings."""
    ods_path = tmp_path / "controls.ods"
    _write_ods(
        ods_path,
        """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
 <office:body>
  <office:spreadsheet>
   <table:table table:name="Sheet1">
    <form:form form:name="MainForm">
     <form:button form:id="btn1" form:name="StartButton" form:label="Start">
      <form:event-listener form:event-name="approveAction"
       xlink:href="vnd.sun.star.script:Module.py$Start?language=Python&amp;location=document"/>
     </form:button>
    </form:form>
   </table:table>
  </office:spreadsheet>
 </office:body>
</office:document-content>
""",
    )

    controls = extract_controls_from_ods(ods_path)
    event_bindings = extract_event_bindings_from_ods(ods_path)

    assert len(controls) == 1
    assert controls[0].id == "btn1"
    assert controls[0].name == "StartButton"
    assert controls[0].control_type == "button"
    assert controls[0].sheet == "Sheet1"
    assert len(event_bindings) == 1
    assert event_bindings[0].control_id == "btn1"
    assert event_bindings[0].target_script_uri is not None


def test_extract_controls_handles_bad_zip(tmp_path: Path) -> None:
    """ODS parse failures should return an empty inventory instead of crashing."""
    ods_path = tmp_path / "bad.ods"
    ods_path.write_text("not a zip")

    assert extract_controls_from_ods(ods_path) == []
    assert extract_event_bindings_from_ods(ods_path) == []


def test_event_binding_resolves_control_when_listener_is_nested(tmp_path: Path) -> None:
    """A listener nested under <office:event-listeners> still resolves to its control."""
    ods_path = tmp_path / "nested.ods"
    _write_ods(
        ods_path,
        """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
 xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
 <office:body>
  <office:spreadsheet>
   <table:table table:name="Sheet1">
    <form:form form:name="MainForm">
     <form:button form:id="btn1" form:name="StartButton" form:label="Start">
      <office:event-listeners>
       <script:event-listener script:event-name="dom:mousedown"
        xlink:href="vnd.sun.star.script:Module.py$Start?language=Python&amp;location=document"/>
      </office:event-listeners>
     </form:button>
    </form:form>
   </table:table>
  </office:spreadsheet>
 </office:body>
</office:document-content>
""",
    )

    bindings = extract_event_bindings_from_ods(ods_path)

    assert len(bindings) == 1
    assert bindings[0].control_id == "btn1"
    assert bindings[0].target_script_uri is not None
