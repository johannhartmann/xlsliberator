"""Transactional macro package update tests."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

import pytest

from xlsliberator.embed_macros import (
    MacroEmbedError,
    embed_python_macros,
    remove_python_macros,
)
from xlsliberator.validation_models import EventBindingIR, SourceRef

MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


def _write_package(path: Path, *, content: str | None = None) -> None:
    manifest = f'''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="{MIMETYPE}"/>
  <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
  <manifest:file-entry manifest:full-path="Scripts/python/" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Scripts/python/keep.py" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Scripts/python/mapped.py" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Scripts/python/replace.py" manifest:media-type="application/binary"/>
  <manifest:file-entry manifest:full-path="Unknown/member.bin" manifest:media-type="application/octet-stream"/>
</manifest:manifest>'''
    if content is None:
        content = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
  <office:scripts/>
</office:document-content>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        archive.writestr("content.xml", content)
        archive.writestr("Scripts/python/keep.py", "def keep():\n    return 'keep'\n")
        archive.writestr(
            "Scripts/python/mapped.py",
            "# xlsliberator-source: VBAProject.Module1.Main\ndef mapped():\n    pass\n",
        )
        archive.writestr("Scripts/python/replace.py", "def old():\n    return 'old'\n")
        archive.writestr("Unknown/member.bin", b"opaque")
        archive.writestr("META-INF/manifest.xml", manifest)


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_upsert_preserves_unrelated_scripts_members_and_manifest_entries(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    _write_package(path)

    embed_python_macros(path, {"replace.py": "def repaired():\n    return 'new'\n"})

    with zipfile.ZipFile(path) as archive:
        assert archive.read("Scripts/python/keep.py") == b"def keep():\n    return 'keep'\n"
        assert b"def repaired" in archive.read("Scripts/python/replace.py")
        assert archive.read("Unknown/member.bin") == b"opaque"
        assert archive.read("Scripts/python/mapped.py") == (
            b"# xlsliberator-source: VBAProject.Module1.Main\ndef mapped():\n    pass\n"
        )
        manifest = archive.read("META-INF/manifest.xml").decode()
        assert "Unknown/member.bin" in manifest
        assert "Scripts/python/keep.py" in manifest


def test_failed_upsert_retains_original_byte_for_byte(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    _write_package(path)
    before = _sha256(path)

    with pytest.raises(MacroEmbedError, match="Invalid Python module"):
        embed_python_macros(path, {"replace.py": "def broken(:\n"})

    assert _sha256(path) == before


def test_unresolved_binding_retains_original(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    _write_package(path)
    before = _sha256(path)
    binding = EventBindingIR(
        id="missing-binding",
        source_ref=SourceRef(
            source_file="book.xlsm", artifact_type="event_binding", artifact_id="missing-binding"
        ),
        event_name="approveAction",
        source_handler="vnd.sun.star.script:VBAProject.Missing.Start?language=Basic",
        target_script_uri="vnd.sun.star.script:replace.py$repaired?language=Python",
    )

    with pytest.raises(MacroEmbedError, match="could not be rewritten"):
        embed_python_macros(path, {"replace.py": "def repaired():\n    pass\n"}, [binding])

    assert _sha256(path) == before


def test_explicit_removal_does_not_delete_other_modules(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    _write_package(path)

    remove_python_macros(path, ["replace.py"])

    with zipfile.ZipFile(path) as archive:
        assert "Scripts/python/replace.py" not in archive.namelist()
        assert "Scripts/python/keep.py" in archive.namelist()
        manifest = archive.read("META-INF/manifest.xml").decode()
        assert "Scripts/python/replace.py" not in manifest
        assert "Scripts/python/keep.py" in manifest


def test_duplicate_casefolded_module_names_are_rejected(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    _write_package(path)

    with pytest.raises(MacroEmbedError, match="Duplicate"):
        remove_python_macros(path, ["Module.py", "module.py"])


def test_non_ascii_module_and_event_target_round_trip(tmp_path: Path) -> None:
    path = tmp_path / "book.ods"
    source = "vnd.sun.star.script:VBAProject.Modul.Start?language=Basic&location=document"
    target = "vnd.sun.star.script:módulo.py$ausführen?language=Python&location=document"
    content = f'''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"><office:scripts xlink:href="{source.replace("&", "&amp;")}"/></office:document-content>'''
    _write_package(path, content=content)
    binding = EventBindingIR(
        id="unicode-binding",
        source_ref=SourceRef(
            source_file="book.xlsm", artifact_type="event_binding", artifact_id="unicode-binding"
        ),
        event_name="approveAction",
        source_handler=source,
        target_script_uri=target,
    )

    embed_python_macros(path, {"módulo.py": "def ausführen():\n    return 'ok'\n"}, [binding])

    with zipfile.ZipFile(path) as archive:
        assert "Scripts/python/módulo.py" in archive.namelist()
        assert "módulo.py$ausführen" in archive.read("content.xml").decode()
