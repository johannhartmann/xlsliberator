"""Contract tests for the native-control ODS seed."""

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_STORED, ZipFile

import pytest

from xlsliberator.native_control_fods import (
    NativeButton,
    NativeSheet,
    inject_native_buttons,
    write_native_button_seed,
)


def test_seed_uses_native_sheet_forms_and_cell_anchored_controls(tmp_path: Path) -> None:
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
    assert 'form:command-type="table"' in serialized
    assert 'form:apply-filter="true"' in serialized
    assert "com.sun.star.form.control.CommandButton" in serialized
    assert "<table:shapes>" not in serialized
    assert root.find(".//table:table-cell/draw:control", namespaces) is not None


def test_seed_rejects_an_unusable_sheet_set(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="at least one sheet"):
        write_native_button_seed(tmp_path / "empty.ods", ())
    with pytest.raises(ValueError, match="visible sheet"):
        write_native_button_seed(
            tmp_path / "hidden.ods",
            (NativeSheet(name="_state", hidden=True),),
        )


def test_injection_preserves_package_and_adds_tagged_native_model(tmp_path: Path) -> None:
    seed = tmp_path / "libreoffice-base.ods"
    write_native_button_seed(seed, (NativeSheet(name="Sheet1"),))

    inject_native_buttons(
        seed,
        (
            NativeSheet(
                name="Sheet1",
                buttons=(
                    NativeButton(
                        name="CertificationButton",
                        label="Run",
                        tag="GameStart",
                        x=1_000,
                        y=1_000,
                        width=5_000,
                    ),
                ),
            ),
        ),
    )

    with ZipFile(seed) as package:
        assert package.infolist()[0].filename == "mimetype"
        assert package.infolist()[0].compress_type == ZIP_STORED
        assert package.read("styles.xml")
        root = ElementTree.fromstring(package.read("content.xml"))
    namespaces = {
        "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
        "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
        "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    }
    button = root.find(".//form:button", namespaces)
    shape = root.find(".//draw:control", namespaces)
    tag = root.find(
        ".//form:property[@form:property-name='Tag']",
        namespaces,
    )
    default_control = root.find(
        ".//form:property[@form:property-name='DefaultControl']",
        namespaces,
    )
    assert button is not None
    assert shape is not None
    assert tag is not None
    assert default_control is not None
    assert default_control.attrib[
        "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}string-value"
    ] == "com.sun.star.form.control.CommandButton"
    assert tag.attrib["{urn:oasis:names:tc:opendocument:xmlns:office:1.0}string-value"] == (
        "GameStart"
    )


def test_injection_rejects_missing_sheet_and_duplicate_forms(tmp_path: Path) -> None:
    seed = tmp_path / "base.ods"
    button = NativeButton(
        name="Button",
        label="Run",
        tag="Run",
        x=0,
        y=0,
        width=1_000,
    )
    write_native_button_seed(seed, (NativeSheet(name="Sheet1"),))
    with pytest.raises(ValueError, match="sheet is missing"):
        inject_native_buttons(seed, (NativeSheet(name="Missing", buttons=(button,)),))

    inject_native_buttons(seed, (NativeSheet(name="Sheet1", buttons=(button,)),))
    with pytest.raises(ValueError, match="already contains forms"):
        inject_native_buttons(seed, (NativeSheet(name="Sheet1", buttons=(button,)),))
