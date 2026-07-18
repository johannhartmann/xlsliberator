"""Adversarial and preservation tests for transactional ODS package edits."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path
from typing import Any

import pytest
from click.testing import CliRunner

from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.odstool import (
    EventBindingSpec,
    OdsPreconditionError,
    OdsToolError,
    bind_event,
    cli,
    diff_packages,
    remove_scripts,
    snapshot_package,
    unbind_event,
    upsert_scripts,
    verify_package,
)

MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"


def _manifest(paths: list[tuple[str, str]]) -> str:
    entries = "\n".join(
        f'  <manifest:file-entry manifest:full-path="{path}" manifest:media-type="{media_type}"/>'
        for path, media_type in paths
    )
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.3">
{entries}
</manifest:manifest>'''


def _content() -> str:
    return """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
 xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"
 xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink">
 <office:body>
  <office:spreadsheet>
   <form:button form:id="run-control" form:name="Run"/>
  </office:spreadsheet>
 </office:body>
</office:document-content>"""


def _write_package(
    path: Path,
    *,
    malformed_content: bool = False,
    signature: bool = False,
) -> Path:
    members: dict[str, bytes | str] = {
        "content.xml": "<broken" if malformed_content else _content(),
        "styles.xml": "<styles/>",
        "Scripts/python/keep.py": "def keep():\n    return 'keep'\n",
        "Scripts/python/repair.py": "def old():\n    return 'old'\n",
        "Unknown/naïve.bin": b"\x00opaque\xff",
    }
    if signature:
        members["META-INF/documentsignatures.xml"] = "<signatures/>"
    manifest_paths = [
        ("/", MIMETYPE),
        ("content.xml", "text/xml"),
        ("styles.xml", "text/xml"),
        ("Scripts/python/", "application/binary"),
        ("Scripts/python/keep.py", "application/binary"),
        ("Scripts/python/repair.py", "application/binary"),
        ("Unknown/naïve.bin", "application/octet-stream"),
    ]
    if signature:
        manifest_paths.append(("META-INF/documentsignatures.xml", "text/xml"))
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        for name, payload in members.items():
            archive.writestr(name, payload)
        archive.writestr("META-INF/manifest.xml", _manifest(manifest_paths))
    return path


def _hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def test_verify_and_cli_expose_complete_package_contract(tmp_path: Path) -> None:
    package = _write_package(tmp_path / "book.ods")

    report = verify_package(package)
    help_result = CliRunner().invoke(cli, ["--help"])
    verify_result = CliRunner().invoke(cli, ["verify", str(package)])

    assert report.valid
    assert [script.module for script in report.scripts] == ["keep.py", "repair.py"]
    assert report.scripts[0].exported_functions == ["keep"]
    assert report.entries[0].path == "mimetype"
    assert report.entries[0].compression == "stored"
    assert {
        "list",
        "verify",
        "inspect-scripts",
        "upsert-script",
        "remove-script",
        "bind-event",
        "unbind-event",
        "diff",
        "snapshot",
    } <= set(help_result.output.split())
    assert verify_result.exit_code == 0
    assert json.loads(verify_result.output)["valid"] is True


def test_upsert_dry_run_precondition_and_commit_preserve_unrelated_data(
    tmp_path: Path,
) -> None:
    package = _write_package(tmp_path / "book.ods", signature=True)
    before = package.read_bytes()
    before_sha = _hash(package)
    with zipfile.ZipFile(package) as archive:
        manifest_before = archive.read("META-INF/manifest.xml")
        unknown_info_before = archive.getinfo("Unknown/naïve.bin")
    source = "def repaired():\n    return 'repaired'\n"

    plan = upsert_scripts(
        package,
        {"répair.py": source},
        expect_sha256=before_sha,
        dry_run=True,
    )

    assert not plan.committed
    assert package.read_bytes() == before
    assert plan.diff.added == ["Scripts/python/répair.py"]
    assert plan.signatures_invalidated

    result = upsert_scripts(
        package,
        {"repair.py": source},
        expect_sha256=before_sha,
    )

    assert result.committed
    assert result.diff.modified == ["Scripts/python/repair.py"]
    assert result.signatures_invalidated
    with zipfile.ZipFile(package) as archive:
        assert archive.read("Unknown/naïve.bin") == b"\x00opaque\xff"
        assert archive.read("Scripts/python/keep.py") == (b"def keep():\n    return 'keep'\n")
        assert archive.read("Scripts/python/repair.py") == source.encode()
        assert archive.read("META-INF/documentsignatures.xml") == b"<signatures/>"
        assert archive.read("META-INF/manifest.xml") == manifest_before
        unknown_info_after = archive.getinfo("Unknown/naïve.bin")
        assert unknown_info_after.date_time == unknown_info_before.date_time
        assert unknown_info_after.external_attr == unknown_info_before.external_attr
        assert unknown_info_after.compress_type == unknown_info_before.compress_type

    with pytest.raises(OdsPreconditionError, match="precondition failed"):
        upsert_scripts(package, {"repair.py": source}, expect_sha256="0" * 64)


def test_remove_one_script_preserves_multiple_other_scripts(tmp_path: Path) -> None:
    package = _write_package(tmp_path / "book.ods")

    result = remove_scripts(package, ["repair"])

    assert result.diff.removed == ["Scripts/python/repair.py"]
    with zipfile.ZipFile(package) as archive:
        assert "Scripts/python/repair.py" not in archive.namelist()
        assert archive.read("Scripts/python/keep.py").startswith(b"def keep")
        manifest = archive.read("META-INF/manifest.xml").decode()
        assert "Scripts/python/repair.py" not in manifest
        assert "Unknown/naïve.bin" in manifest


def test_bind_and_unbind_event_validate_control_module_and_export(tmp_path: Path) -> None:
    package = _write_package(tmp_path / "book.ods")
    binding = EventBindingSpec(
        id="run-click",
        control_id="run-control",
        event_name="dom:click",
        module="keep.py",
        function="keep",
    )

    bound = bind_event(package, binding)

    assert bound.diff.modified == ["content.xml"]
    report = verify_package(package)
    assert report.valid
    with zipfile.ZipFile(package) as archive:
        content = archive.read("content.xml").decode()
        assert "run-click" in content
        assert "keep.py$keep" in content

    unbound = unbind_event(package, "run-click")

    assert unbound.diff.modified == ["content.xml"]
    with zipfile.ZipFile(package) as archive:
        assert "run-click" not in archive.read("content.xml").decode()


@pytest.mark.parametrize(
    "binding",
    [
        EventBindingSpec(
            id="missing-control",
            control_id="absent",
            event_name="dom:click",
            module="keep.py",
            function="keep",
        ),
        EventBindingSpec(
            id="missing-export",
            control_id="run-control",
            event_name="dom:click",
            module="keep.py",
            function="absent",
        ),
    ],
)
def test_unresolved_event_binding_rolls_back(
    tmp_path: Path,
    binding: EventBindingSpec,
) -> None:
    package = _write_package(tmp_path / "book.ods")
    before = package.read_bytes()

    with pytest.raises(OdsToolError):
        bind_event(package, binding)

    assert package.read_bytes() == before


def test_duplicate_paths_malformed_xml_and_path_traversal_fail_closed(
    tmp_path: Path,
) -> None:
    duplicate = tmp_path / "duplicate.ods"
    with zipfile.ZipFile(duplicate, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        archive.writestr("content.xml", _content())
        with pytest.warns(UserWarning, match="Duplicate name"):
            archive.writestr("content.xml", _content())
        archive.writestr(
            "META-INF/manifest.xml",
            _manifest([("/", MIMETYPE), ("content.xml", "text/xml")]),
        )
    malformed = _write_package(tmp_path / "malformed.ods", malformed_content=True)
    traversal = tmp_path / "traversal.ods"
    with zipfile.ZipFile(traversal, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        archive.writestr("../escape", b"bad")
        archive.writestr("content.xml", _content())
        archive.writestr(
            "META-INF/manifest.xml",
            _manifest([("/", MIMETYPE), ("content.xml", "text/xml")]),
        )

    assert "duplicate member paths" in "; ".join(verify_package(duplicate).errors)
    assert "Malformed XML" in "; ".join(verify_package(malformed).errors)
    assert "unsafe member path" in "; ".join(verify_package(traversal).errors)


def test_malformed_manifest_entity_and_compressed_mimetype_fail_closed(tmp_path: Path) -> None:
    malformed_manifest = _write_package(tmp_path / "malformed-manifest.ods")
    with zipfile.ZipFile(malformed_manifest, "a") as archive:
        with pytest.warns(UserWarning, match="Duplicate name"):
            archive.writestr("META-INF/manifest.xml", "<manifest:manifest")

    entity_manifest = tmp_path / "entity-manifest.ods"
    entity_xml = f"""<?xml version="1.0"?>
<!DOCTYPE manifest [<!ENTITY leak SYSTEM "file:///etc/passwd">]>
<manifest:manifest xmlns:manifest="{MANIFEST_NS}">
 <manifest:file-entry manifest:full-path="/" manifest:media-type="&leak;"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
</manifest:manifest>"""
    with zipfile.ZipFile(entity_manifest, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_STORED)
        archive.writestr("content.xml", _content())
        archive.writestr("META-INF/manifest.xml", entity_xml)

    compressed_mimetype = tmp_path / "compressed-mimetype.ods"
    with zipfile.ZipFile(compressed_mimetype, "w") as archive:
        archive.writestr("mimetype", MIMETYPE, compress_type=zipfile.ZIP_DEFLATED)
        archive.writestr("content.xml", _content())
        archive.writestr(
            "META-INF/manifest.xml",
            _manifest([("/", MIMETYPE), ("content.xml", "text/xml")]),
        )

    assert not verify_package(malformed_manifest).valid
    assert "Malformed XML" in "; ".join(verify_package(entity_manifest).errors)
    assert "stored without compression" in "; ".join(verify_package(compressed_mimetype).errors)


def test_failed_partial_candidate_write_leaves_original_byte_for_byte(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    package = _write_package(tmp_path / "book.ods")
    before = package.read_bytes()

    def fail_after_partial_write(
        _archive: zipfile.ZipFile,
        candidate: Path,
        _replacements: dict[str, bytes | None],
    ) -> None:
        candidate.write_bytes(b"partial")
        raise OSError("simulated disk failure")

    monkeypatch.setattr("xlsliberator.odstool._write_candidate", fail_after_partial_write)

    with pytest.raises(OSError, match="simulated disk failure"):
        upsert_scripts(package, {"repair.py": "def fixed():\n    pass\n"})

    assert package.read_bytes() == before
    assert not list(tmp_path.glob("*.odstool.tmp"))


def test_concurrent_package_change_is_not_overwritten(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    package = _write_package(tmp_path / "book.ods")
    external = _write_package(tmp_path / "external.ods", signature=True).read_bytes()

    from xlsliberator import odstool

    real_write_candidate = odstool._write_candidate

    def change_original_after_candidate(
        archive: zipfile.ZipFile,
        candidate: Path,
        replacements: dict[str, bytes | None],
    ) -> None:
        real_write_candidate(archive, candidate, replacements)
        package.write_bytes(external)

    monkeypatch.setattr(odstool, "_write_candidate", change_original_after_candidate)

    with pytest.raises(OdsPreconditionError, match="changed during mutation"):
        upsert_scripts(package, {"repair.py": "def fixed():\n    return 1\n"})

    assert package.read_bytes() == external
    assert not list(tmp_path.glob("*.odstool.tmp"))


def test_atomic_replace_preserves_package_permissions(tmp_path: Path) -> None:
    package = _write_package(tmp_path / "book.ods")
    package.chmod(0o640)

    upsert_scripts(package, {"repair.py": "def fixed():\n    return 1\n"})

    assert package.stat().st_mode & 0o777 == 0o640


def test_namespace_edit_preserves_xml_comments(tmp_path: Path) -> None:
    package = _write_package(tmp_path / "book.ods")
    with zipfile.ZipFile(package) as archive:
        members = {info.filename: archive.read(info) for info in archive.infolist()}
        infos = {info.filename: info for info in archive.infolist()}
    members["content.xml"] = members["content.xml"].replace(
        b"<office:body>",
        b"<office:body><!--keep-this-comment-->",
    )
    with zipfile.ZipFile(package, "w") as archive:
        for name, payload in members.items():
            archive.writestr(infos[name], payload)

    bind_event(
        package,
        EventBindingSpec(
            id="comment-test",
            control_id="run-control",
            event_name="dom:click",
            module="keep.py",
            function="keep",
        ),
    )

    with zipfile.ZipFile(package) as archive:
        assert b"keep-this-comment" in archive.read("content.xml")


def test_snapshot_and_diff_preserve_raw_unknown_parts(tmp_path: Path) -> None:
    before = _write_package(tmp_path / "before.ods")
    after = tmp_path / "after.ods"
    after.write_bytes(before.read_bytes())
    upsert_scripts(after, {"repair.py": "def fixed():\n    return 1\n"})

    diff = diff_packages(before, after)
    destination = snapshot_package(after, tmp_path / "snapshot")

    assert diff.modified == ["Scripts/python/repair.py"]
    assert (destination / "raw/Unknown/naïve.bin").read_bytes() == b"\x00opaque\xff"
    assert json.loads((destination / "summary.json").read_text())["valid"] is True


def test_existing_embed_api_delegates_to_transactional_layer(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    package = _write_package(tmp_path / "book.ods")
    calls: list[tuple[Path, dict[str, str]]] = []
    real_upsert = upsert_scripts

    def recording_upsert(
        path: str | Path,
        modules: dict[str, str],
        **kwargs: Any,
    ) -> Any:
        calls.append((Path(path), modules))
        return real_upsert(path, modules, **kwargs)

    monkeypatch.setattr("xlsliberator.odstool.upsert_scripts", recording_upsert)

    embed_python_macros(package, {"repair.py": "def repaired():\n    return 1\n"})

    assert calls == [(package, {"repair.py": "def repaired():\n    return 1\n"})]
