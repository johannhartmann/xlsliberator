"""Contract tests for the native-control ODS seed."""

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_STORED, ZipFile

import pytest

from xlsliberator.native_control_fods import (
    NativeButton,
    NativeSheet,
    write_native_button_seed,
)


def test_seed_uses_sheet_local_forms_and_stable_control_references(tmp_path: Path) -> None:
    seed = tmp_path / "controls.ods"

    write_native_button_seed(
        seed,
        (
            NativeSheet(
                name='game "certification"',
                buttons=(
                    NativeButton(
                        name='Certification "Button"',
                        label="Start & play <now>",
                        tag='GameStart "safe"',
                        x=1_000,
                        y=2_000,
                        width=5_000,
                    ),
                ),
            ),
            NativeSheet(name="_state", hidden=True),
        ),
    )

    with ZipFile(seed) as package:
        infos = package.infolist()
        assert infos[0].filename == "mimetype"
        assert infos[0].compress_type == ZIP_STORED
        assert package.read("mimetype") == b"application/vnd.oasis.opendocument.spreadsheet"
        assert {"META-INF/manifest.xml", "content.xml", "styles.xml"} <= set(package.namelist())
        content = package.read("content.xml")
    root = ElementTree.fromstring(content)
    namespaces = {
        "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
        "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    }
    tables = root.findall(".//table:table", namespaces)
    assert len(tables) == 2
    assert tables[0].find("./office:forms/form:form/form:button", namespaces) is not None
    assert tables[1].find("./office:forms", namespaces) is None

    button = root.find(".//form:button", namespaces)
    shape = root.find(".//draw:control", namespaces)
    assert button is not None
    assert shape is not None
    form_id = button.attrib["{urn:oasis:names:tc:opendocument:xmlns:form:1.0}id"]
    assert shape.attrib["{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}control"] == form_id
    serialized = content.decode("utf-8")
    assert "Start &amp; play &lt;now&gt;" in serialized
    assert 'table:name="game &quot;certification&quot;"' in serialized
    assert 'form:name="Certification &quot;Button&quot;"' in serialized
    assert 'form:delay-for-repeat="PT0.050000000S"' in serialized
    assert 'xlink:href=""' in serialized
    assert "<table:shapes>" in serialized


def test_seed_rejects_an_unusable_sheet_set(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="at least one sheet"):
        write_native_button_seed(tmp_path / "empty.ods", ())
    with pytest.raises(ValueError, match="visible sheet"):
        write_native_button_seed(
            tmp_path / "hidden.ods",
            (NativeSheet(name="_state", hidden=True),),
        )
