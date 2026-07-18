"""Deterministic, transactional editing for untrusted ODS packages."""

from __future__ import annotations

import ast
import hashlib
import json
import os
import secrets
import shutil
import tempfile

# Construction and serialization only; untrusted XML is parsed with defusedxml.
import xml.etree.ElementTree as ET  # nosec B405
import zipfile
from collections.abc import Callable, Mapping
from pathlib import Path, PurePosixPath
from typing import Any, cast
from urllib.parse import parse_qs

import click
import yaml
from defusedxml.common import DefusedXmlException
from defusedxml.ElementTree import DefusedXMLParser
from pydantic import BaseModel, ConfigDict, Field

from xlsliberator.container_boundary import ContainerBoundaryError, require_application_container
from xlsliberator.validation_models import EventBindingIR

MIMETYPE = b"application/vnd.oasis.opendocument.spreadsheet"
MANIFEST_PATH = "META-INF/manifest.xml"
CONTENT_PATH = "content.xml"
SCRIPT_ROOT = "Scripts/python/"
MANIFEST_NS = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
SCRIPT_NS = "urn:oasis:names:tc:opendocument:xmlns:script:1.0"
XLINK_NS = "http://www.w3.org/1999/xlink"
BINDING_NS = "urn:xlsliberator:event-bindings:1.0"
SCHEMA_VERSION = "1.0"

ET.register_namespace("manifest", MANIFEST_NS)
ET.register_namespace("script", SCRIPT_NS)
ET.register_namespace("xlink", XLINK_NS)
ET.register_namespace("xlsliberator", BINDING_NS)


class OdsToolError(RuntimeError):
    """Base error for deterministic ODS package operations."""


class OdsPreconditionError(OdsToolError):
    """Raised when a caller's content precondition does not match."""


class PackageEntry(BaseModel):
    """One verified ODS ZIP member."""

    model_config = ConfigDict(extra="forbid")

    path: str
    size: int
    compressed_size: int
    compression: str
    crc32: str
    sha256: str


class ScriptRecord(BaseModel):
    """One embedded Python module and its exported top-level functions."""

    model_config = ConfigDict(extra="forbid")

    module: str
    package_path: str
    size: int
    sha256: str
    exported_functions: list[str] = Field(default_factory=list)


class PackageVerification(BaseModel):
    """Fail-closed structural and script-target verification."""

    model_config = ConfigDict(extra="forbid")

    schema_version: str = SCHEMA_VERSION
    path: str
    sha256: str | None = None
    valid: bool
    entries: list[PackageEntry] = Field(default_factory=list)
    scripts: list[ScriptRecord] = Field(default_factory=list)
    signature_entries: list[str] = Field(default_factory=list)
    errors: list[str] = Field(default_factory=list)
    warnings: list[str] = Field(default_factory=list)


class PackageDiff(BaseModel):
    """Deterministic member-level before/after package diff."""

    model_config = ConfigDict(extra="forbid")

    before_sha256: str
    after_sha256: str
    added: list[str] = Field(default_factory=list)
    removed: list[str] = Field(default_factory=list)
    modified: list[str] = Field(default_factory=list)
    unchanged: int = 0
    signatures_before: list[str] = Field(default_factory=list)
    signatures_after: list[str] = Field(default_factory=list)


class MutationResult(BaseModel):
    """Evidence for a validated package mutation or dry-run plan."""

    model_config = ConfigDict(extra="forbid")

    schema_version: str = SCHEMA_VERSION
    operation: str
    package_path: str
    dry_run: bool
    committed: bool
    before_sha256: str
    after_sha256: str
    diff: PackageDiff
    signatures_invalidated: bool = False
    warnings: list[str] = Field(default_factory=list)


class EventBindingSpec(BaseModel):
    """Namespace-aware binding request loaded from YAML."""

    model_config = ConfigDict(extra="forbid")

    id: str = Field(min_length=1, max_length=200)
    control_id: str = Field(min_length=1, max_length=500)
    event_name: str = Field(min_length=1, max_length=200)
    module: str = Field(min_length=1, max_length=255)
    function: str = Field(min_length=1, max_length=255)

    @property
    def target_script_uri(self) -> str:
        """Return the canonical document-local LibreOffice Python URI."""
        module = _normalize_module_name(self.module)
        if not self.function.isidentifier():
            raise OdsToolError(f"Invalid Python function name: {self.function!r}")
        return f"vnd.sun.star.script:{module}${self.function}?language=Python&location=document"


ReplacementMap = Mapping[str, bytes | None]
MutationBuilder = Callable[[zipfile.ZipFile], ReplacementMap]


def list_package(path: str | Path) -> PackageVerification:
    """Return a verified member inventory without changing the package."""
    return verify_package(path)


def verify_package(path: str | Path) -> PackageVerification:
    """Verify ODS ZIP, XML, manifest, scripts, and Python event targets."""
    require_application_container()
    package = Path(path)
    errors: list[str] = []
    warnings: list[str] = []
    entries: list[PackageEntry] = []
    scripts: list[ScriptRecord] = []
    signatures: list[str] = []
    package_sha: str | None = None

    if not package.is_file():
        errors.append("ODS package does not exist")
        return PackageVerification(path=str(package), valid=False, errors=errors)
    package_sha = _hash_path(package)
    if not zipfile.is_zipfile(package):
        errors.append("ODS package is not a ZIP archive")
        return PackageVerification(
            path=str(package),
            sha256=package_sha,
            valid=False,
            errors=errors,
        )

    try:
        with zipfile.ZipFile(package) as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            _verify_member_names(infos, errors)
            _verify_mimetype(archive, infos, errors)
            corrupt = archive.testzip()
            if corrupt is not None:
                errors.append(f"ODS member failed CRC validation: {corrupt}")
            for required in (CONTENT_PATH, MANIFEST_PATH):
                if required not in names:
                    errors.append(f"ODS package is missing {required}")

            signatures = sorted(name for name in names if _is_signature_entry(name))
            if signatures:
                warnings.append(
                    "Package contains digital signatures; any content mutation invalidates them"
                )

            if not errors:
                entries = [_entry_record(archive, info) for info in infos]
                manifest_root = _parse_required_xml(archive, MANIFEST_PATH, errors)
                content_root = _parse_required_xml(archive, CONTENT_PATH, errors)
                if manifest_root is not None:
                    manifest_paths = _manifest_paths(manifest_root, errors)
                    scripts = _inspect_scripts_from_archive(archive, manifest_paths, errors)
                else:
                    manifest_paths = set()
                if content_root is not None:
                    _verify_python_event_targets(content_root, scripts, errors)
                _verify_manifest_members(names, manifest_paths, errors)
    except (OSError, zipfile.BadZipFile, KeyError, RuntimeError) as exc:
        errors.append(f"Invalid ODS package: {exc}")

    return PackageVerification(
        path=str(package),
        sha256=package_sha,
        valid=not errors,
        entries=entries,
        scripts=scripts,
        signature_entries=signatures,
        errors=errors,
        warnings=warnings,
    )


def inspect_scripts(path: str | Path) -> PackageVerification:
    """Return verified embedded-script metadata."""
    return verify_package(path)


def upsert_scripts(
    path: str | Path,
    modules: Mapping[str, str],
    *,
    event_bindings: list[EventBindingIR] | None = None,
    expect_sha256: str | None = None,
    dry_run: bool = False,
) -> MutationResult:
    """Upsert only named scripts and preserve every unrelated package member."""
    normalized = _normalize_modules(modules)
    if not normalized:
        raise OdsToolError("At least one Python module is required")

    def build(archive: zipfile.ZipFile) -> ReplacementMap:
        manifest = _updated_manifest(
            archive,
            add_modules=list(normalized),
            remove_modules=[],
        )
        replacements: dict[str, bytes | None] = {
            MANIFEST_PATH: manifest,
            **{
                f"{SCRIPT_ROOT}{module}": source.encode("utf-8")
                for module, source in normalized.items()
            },
        }
        if event_bindings:
            from xlsliberator.event_binding_writer import rewrite_event_bindings

            content = archive.read(CONTENT_PATH).decode("utf-8")
            rewritten, unresolved = rewrite_event_bindings(
                content,
                dict(normalized),
                event_bindings,
            )
            if unresolved:
                identifiers = ", ".join(binding.id for binding in unresolved)
                raise OdsToolError(f"Event binding(s) could not be rewritten: {identifiers}")
            replacements[CONTENT_PATH] = rewritten.encode("utf-8")
        if replacements[MANIFEST_PATH] == archive.read(MANIFEST_PATH):
            del replacements[MANIFEST_PATH]
        return replacements

    return _mutate_package(
        Path(path),
        operation="upsert-script",
        build_replacements=build,
        expect_sha256=expect_sha256,
        dry_run=dry_run,
    )


def remove_scripts(
    path: str | Path,
    modules: list[str],
    *,
    expect_sha256: str | None = None,
    dry_run: bool = False,
) -> MutationResult:
    """Remove only explicitly named Python modules."""
    normalized = _normalize_module_names(modules)
    if not normalized:
        raise OdsToolError("At least one Python module is required")

    def build(archive: zipfile.ZipFile) -> ReplacementMap:
        names = set(archive.namelist())
        missing = [name for name in normalized if f"{SCRIPT_ROOT}{name}" not in names]
        if missing:
            raise OdsToolError(f"Scripts do not exist: {', '.join(missing)}")
        replacements: dict[str, bytes | None] = {
            f"{SCRIPT_ROOT}{name}": None for name in normalized
        }
        replacements[MANIFEST_PATH] = _updated_manifest(
            archive,
            add_modules=[],
            remove_modules=normalized,
        )
        return replacements

    return _mutate_package(
        Path(path),
        operation="remove-script",
        build_replacements=build,
        expect_sha256=expect_sha256,
        dry_run=dry_run,
    )


def bind_event(
    path: str | Path,
    binding: EventBindingSpec,
    *,
    expect_sha256: str | None = None,
    dry_run: bool = False,
) -> MutationResult:
    """Bind one validated event target to a uniquely identified control."""

    def build(archive: zipfile.ZipFile) -> ReplacementMap:
        scripts = _script_exports_for_mutation(archive)
        module = _normalize_module_name(binding.module)
        if binding.function not in scripts.get(module, set()):
            raise OdsToolError(
                f"Event target does not resolve to an exported function: "
                f"{module}${binding.function}"
            )
        root = _safe_xml_root(archive.read(CONTENT_PATH), CONTENT_PATH)
        binding_attr = f"{{{BINDING_NS}}}binding-id"
        if any(element.get(binding_attr) == binding.id for element in root.iter()):
            raise OdsToolError(f"Event binding already exists: {binding.id}")
        controls = [
            element
            for element in root.iter()
            if binding.control_id
            in {
                value for key, value in element.attrib.items() if _local_name(key) in {"id", "name"}
            }
        ]
        if len(controls) != 1:
            raise OdsToolError(
                f"Control locator {binding.control_id!r} resolved to {len(controls)} elements"
            )
        listener = ET.SubElement(controls[0], f"{{{SCRIPT_NS}}}event-listener")
        listener.set(f"{{{SCRIPT_NS}}}event-name", binding.event_name)
        listener.set(f"{{{XLINK_NS}}}href", binding.target_script_uri)
        listener.set(binding_attr, binding.id)
        return {CONTENT_PATH: _serialize_xml(root)}

    return _mutate_package(
        Path(path),
        operation="bind-event",
        build_replacements=build,
        expect_sha256=expect_sha256,
        dry_run=dry_run,
    )


def unbind_event(
    path: str | Path,
    binding_id: str,
    *,
    expect_sha256: str | None = None,
    dry_run: bool = False,
) -> MutationResult:
    """Remove exactly one binding created by :func:`bind_event`."""
    if not binding_id:
        raise OdsToolError("Binding ID must not be empty")

    def build(archive: zipfile.ZipFile) -> ReplacementMap:
        root = _safe_xml_root(archive.read(CONTENT_PATH), CONTENT_PATH)
        binding_attr = f"{{{BINDING_NS}}}binding-id"
        parents = {child: parent for parent in root.iter() for child in parent}
        matches = [element for element in root.iter() if element.get(binding_attr) == binding_id]
        if len(matches) != 1:
            raise OdsToolError(f"Binding ID {binding_id!r} resolved to {len(matches)} elements")
        parent = parents.get(matches[0])
        if parent is None:
            raise OdsToolError(f"Binding {binding_id!r} has no removable parent")
        parent.remove(matches[0])
        return {CONTENT_PATH: _serialize_xml(root)}

    return _mutate_package(
        Path(path),
        operation="unbind-event",
        build_replacements=build,
        expect_sha256=expect_sha256,
        dry_run=dry_run,
    )


def diff_packages(before: str | Path, after: str | Path) -> PackageDiff:
    """Compare two verified ODS packages by member payload."""
    require_application_container()
    before_path = Path(before)
    after_path = Path(after)
    before_verification = verify_package(before_path)
    after_verification = verify_package(after_path)
    _require_valid(before_verification)
    _require_valid(after_verification)
    return _diff_verified(before_path, after_path, before_verification, after_verification)


def snapshot_package(path: str | Path, output: str | Path) -> Path:
    """Create a transactional, model-readable package snapshot."""
    require_application_container()
    package = Path(path)
    verification = verify_package(package)
    _require_valid(verification)
    destination = Path(output)
    if destination.exists() and (not destination.is_dir() or any(destination.iterdir())):
        raise OdsToolError(f"Refusing to replace non-empty snapshot output: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = Path(tempfile.mkdtemp(prefix=".odstool-snapshot-", dir=destination.parent))
    try:
        raw_root = temporary / "raw"
        with zipfile.ZipFile(package) as archive:
            for info in archive.infolist():
                _validate_member_path(info.filename)
                if info.is_dir():
                    continue
                target = raw_root.joinpath(*PurePosixPath(info.filename).parts)
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(archive.read(info))
        (temporary / "summary.json").write_text(
            verification.model_dump_json(indent=2) + "\n",
            encoding="utf-8",
        )
        (temporary / "package-tree.txt").write_text(
            "\n".join(
                f"{entry.path}\t{entry.size}\t{entry.sha256}" for entry in verification.entries
            )
            + "\n",
            encoding="utf-8",
        )
        _fsync_tree(temporary)
        if destination.exists():
            destination.rmdir()
        os.replace(temporary, destination)
        _fsync_directory(destination.parent)
    except Exception:
        shutil.rmtree(temporary, ignore_errors=True)
        raise
    return destination


def _mutate_package(
    package: Path,
    *,
    operation: str,
    build_replacements: MutationBuilder,
    expect_sha256: str | None,
    dry_run: bool,
) -> MutationResult:
    require_application_container()
    before = verify_package(package)
    _require_valid(before)
    assert before.sha256 is not None
    if expect_sha256 is not None and not secrets.compare_digest(
        before.sha256.lower(), expect_sha256.lower()
    ):
        raise OdsPreconditionError(
            f"Package SHA-256 precondition failed: expected {expect_sha256}, found {before.sha256}"
        )

    descriptor, raw_temp_path = tempfile.mkstemp(
        prefix=f".{package.name}.",
        suffix=".odstool.tmp",
        dir=package.parent,
    )
    os.close(descriptor)
    candidate: Path | None = Path(raw_temp_path)
    try:
        with zipfile.ZipFile(package) as archive:
            replacements = dict(build_replacements(archive))
            _validate_replacement_paths(replacements)
            assert candidate is not None
            _write_candidate(archive, candidate, replacements)
        assert candidate is not None
        os.chmod(candidate, package.stat().st_mode & 0o7777)
        _fsync_file(candidate)
        after = verify_package(candidate)
        _require_valid(after)
        assert after.sha256 is not None
        diff = _diff_verified(package, candidate, before, after)
        signatures_invalidated = bool(before.signature_entries and _content_changed(diff))
        warnings = list(after.warnings)
        if signatures_invalidated:
            warnings.append(
                "The operation changes signed package content and invalidates existing signatures"
            )
        if not dry_run:
            current_sha = _hash_path(package)
            if not secrets.compare_digest(before.sha256, current_sha):
                raise OdsPreconditionError(
                    "Package changed during mutation: "
                    f"started at {before.sha256}, now {current_sha}"
                )
            os.replace(candidate, package)
            _fsync_directory(package.parent)
            candidate = None
        return MutationResult(
            operation=operation,
            package_path=str(package),
            dry_run=dry_run,
            committed=not dry_run,
            before_sha256=before.sha256,
            after_sha256=after.sha256,
            diff=diff,
            signatures_invalidated=signatures_invalidated,
            warnings=warnings,
        )
    except Exception:
        if candidate is not None:
            candidate.unlink(missing_ok=True)
        raise
    finally:
        if candidate is not None:
            candidate.unlink(missing_ok=True)


def _write_candidate(
    archive: zipfile.ZipFile,
    candidate: Path,
    replacements: dict[str, bytes | None],
) -> None:
    """Write one complete candidate package; isolated for fault-injection tests."""
    seen: set[str] = set()
    with zipfile.ZipFile(candidate, "w") as output:
        for info in archive.infolist():
            name = info.filename
            if name in seen:
                raise OdsToolError(f"Duplicate package member path: {name}")
            seen.add(name)
            if name in replacements:
                payload = replacements.pop(name)
                if payload is not None:
                    output.writestr(info, payload)
                continue
            output.writestr(info, archive.read(info))
        for name, payload in replacements.items():
            if payload is None:
                raise OdsToolError(f"Cannot remove missing package member: {name}")
            output.writestr(name, payload, compress_type=zipfile.ZIP_DEFLATED)


def _verify_member_names(infos: list[zipfile.ZipInfo], errors: list[str]) -> None:
    names = [info.filename for info in infos]
    if len(names) != len(set(names)):
        errors.append("ODS package contains duplicate member paths")
    for name in names:
        try:
            _validate_member_path(name)
        except OdsToolError as exc:
            errors.append(str(exc))


def _verify_mimetype(
    archive: zipfile.ZipFile,
    infos: list[zipfile.ZipInfo],
    errors: list[str],
) -> None:
    if not infos or infos[0].filename != "mimetype":
        errors.append("ODS mimetype entry is missing or not first")
        return
    if infos[0].compress_type != zipfile.ZIP_STORED:
        errors.append("ODS mimetype entry must be stored without compression")
    try:
        if archive.read(infos[0]) != MIMETYPE:
            errors.append("ODS mimetype value is invalid")
    except (KeyError, RuntimeError, zipfile.BadZipFile) as exc:
        errors.append(f"ODS mimetype cannot be read: {exc}")


def _entry_record(archive: zipfile.ZipFile, info: zipfile.ZipInfo) -> PackageEntry:
    payload = archive.read(info)
    return PackageEntry(
        path=info.filename,
        size=info.file_size,
        compressed_size=info.compress_size,
        compression=_compression_name(info.compress_type),
        crc32=f"{info.CRC:08x}",
        sha256=hashlib.sha256(payload).hexdigest(),
    )


def _parse_required_xml(
    archive: zipfile.ZipFile,
    name: str,
    errors: list[str],
) -> ET.Element | None:
    try:
        return _safe_xml_root(archive.read(name), name)
    except (KeyError, OdsToolError) as exc:
        errors.append(str(exc))
        return None


def _safe_xml_root(payload: bytes, name: str) -> ET.Element:
    try:
        parser = DefusedXMLParser(
            target=ET.TreeBuilder(insert_comments=True, insert_pis=True),
            forbid_dtd=True,
            forbid_entities=True,
            forbid_external=True,
        )
        parser.feed(payload)
        return parser.close()
    except (ET.ParseError, DefusedXmlException) as exc:
        raise OdsToolError(f"Malformed XML in {name}: {exc}") from exc


def _manifest_paths(root: ET.Element, errors: list[str]) -> set[str]:
    paths: list[str] = []
    for entry in root.findall(f".//{{{MANIFEST_NS}}}file-entry"):
        path = entry.get(f"{{{MANIFEST_NS}}}full-path")
        if path is None:
            errors.append("Manifest contains a file-entry without full-path")
            continue
        paths.append(path)
    if len(paths) != len(set(paths)):
        errors.append("Manifest contains duplicate full-path entries")
    return set(paths)


def _verify_manifest_members(
    names: list[str],
    manifest_paths: set[str],
    errors: list[str],
) -> None:
    name_set = set(names)
    for name in sorted(name_set):
        if name in {"mimetype", MANIFEST_PATH} or name.endswith("/"):
            continue
        if name not in manifest_paths:
            errors.append(f"Manifest omits package member: {name}")
    dangling = sorted(
        name
        for name in manifest_paths
        if name not in {"/", "mimetype", MANIFEST_PATH}
        and not name.endswith("/")
        and name not in name_set
    )
    if dangling:
        errors.append(f"Manifest references missing package members: {dangling}")


def _inspect_scripts_from_archive(
    archive: zipfile.ZipFile,
    manifest_paths: set[str],
    errors: list[str],
) -> list[ScriptRecord]:
    records: list[ScriptRecord] = []
    for name in sorted(
        member
        for member in archive.namelist()
        if member.startswith(SCRIPT_ROOT) and member.endswith(".py")
    ):
        if name not in manifest_paths:
            continue
        try:
            payload = archive.read(name)
            source = payload.decode("utf-8")
            exported = sorted(_exported_functions(source, name))
            records.append(
                ScriptRecord(
                    module=PurePosixPath(name).name,
                    package_path=name,
                    size=len(payload),
                    sha256=hashlib.sha256(payload).hexdigest(),
                    exported_functions=exported,
                )
            )
        except (UnicodeDecodeError, SyntaxError, ValueError) as exc:
            errors.append(f"Invalid embedded Python module {name}: {exc}")
    return records


def _verify_python_event_targets(
    root: ET.Element,
    scripts: list[ScriptRecord],
    errors: list[str],
) -> None:
    exports = {script.module: set(script.exported_functions) for script in scripts}
    for element in root.iter():
        target = element.get(f"{{{XLINK_NS}}}href")
        if not target or "language=Python" not in target:
            continue
        try:
            module, function = _parse_script_uri(target)
        except OdsToolError as exc:
            errors.append(str(exc))
            continue
        if function not in exports.get(module, set()):
            errors.append(f"Python event target does not resolve: {module}${function}")


def _updated_manifest(
    archive: zipfile.ZipFile,
    *,
    add_modules: list[str],
    remove_modules: list[str],
) -> bytes:
    original = archive.read(MANIFEST_PATH)
    root = _safe_xml_root(original, MANIFEST_PATH)
    owned_paths = {
        *(f"{SCRIPT_ROOT}{name}" for name in add_modules),
        *(f"{SCRIPT_ROOT}{name}" for name in remove_modules),
    }
    changed = False
    existing_counts: dict[str, int] = {}
    for entry in root.findall(f".//{{{MANIFEST_NS}}}file-entry"):
        full_path = entry.get(f"{{{MANIFEST_NS}}}full-path", "")
        existing_counts[full_path] = existing_counts.get(full_path, 0) + 1
    for parent in root.iter():
        for child in list(parent):
            if child.tag != f"{{{MANIFEST_NS}}}file-entry":
                continue
            full_path = child.get(f"{{{MANIFEST_NS}}}full-path", "")
            should_remove = full_path in {f"{SCRIPT_ROOT}{name}" for name in remove_modules} or (
                full_path in {f"{SCRIPT_ROOT}{name}" for name in add_modules}
                and existing_counts.get(full_path, 0) > 1
            )
            if full_path in owned_paths and should_remove:
                parent.remove(child)
                existing_counts[full_path] -= 1
                changed = True
    manifest_paths = {
        entry.get(f"{{{MANIFEST_NS}}}full-path", "")
        for entry in root.findall(f".//{{{MANIFEST_NS}}}file-entry")
    }
    if SCRIPT_ROOT not in manifest_paths:
        _add_manifest_entry(root, SCRIPT_ROOT, "application/binary")
        changed = True
    for module in add_modules:
        script_path = f"{SCRIPT_ROOT}{module}"
        if script_path not in manifest_paths:
            _add_manifest_entry(root, script_path, "application/binary")
            changed = True
    return _serialize_xml(root) if changed else original


def _script_exports_for_mutation(archive: zipfile.ZipFile) -> dict[str, set[str]]:
    errors: list[str] = []
    manifest = _safe_xml_root(archive.read(MANIFEST_PATH), MANIFEST_PATH)
    records = _inspect_scripts_from_archive(archive, _manifest_paths(manifest, errors), errors)
    if errors:
        raise OdsToolError("; ".join(errors))
    return {record.module: set(record.exported_functions) for record in records}


def _normalize_modules(modules: Mapping[str, str]) -> dict[str, str]:
    names = _normalize_module_names(list(modules))
    normalized: dict[str, str] = {}
    for original, name in zip(modules, names, strict=True):
        source = modules[original]
        try:
            _exported_functions(source, name)
        except (SyntaxError, ValueError) as exc:
            raise OdsToolError(f"Invalid Python module {name}: {exc}") from exc
        normalized[name] = source
    return normalized


def _normalize_module_names(modules: list[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for raw in modules:
        name = _normalize_module_name(raw)
        collision = name.casefold()
        if collision in seen:
            raise OdsToolError(f"Duplicate Python module name: {name}")
        seen.add(collision)
        normalized.append(name)
    return normalized


def _normalize_module_name(raw: str) -> str:
    name = raw if raw.endswith(".py") else f"{raw}.py"
    if (
        not name
        or PurePosixPath(name).name != name
        or "/" in name
        or "\\" in name
        or "\x00" in name
    ):
        raise OdsToolError(f"Invalid Python module name: {raw!r}")
    return name


def _exported_functions(source: str, name: str) -> set[str]:
    tree = ast.parse(source, filename=name)
    return {
        node.name for node in tree.body if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }


def _parse_script_uri(target: str) -> tuple[str, str]:
    marker = "vnd.sun.star.script:"
    if not target.startswith(marker) or "$" not in target or "?" not in target:
        raise OdsToolError(f"Invalid Python event target URI: {target!r}")
    module, function_and_query = target[len(marker) :].split("$", 1)
    function, query = function_and_query.split("?", 1)
    parameters = parse_qs(query, keep_blank_values=True)
    if parameters.get("language") != ["Python"] or parameters.get("location") != ["document"]:
        raise OdsToolError(f"Invalid Python event target URI: {target!r}")
    if not function.isidentifier():
        raise OdsToolError(f"Invalid Python event target function: {function!r}")
    return _normalize_module_name(module), function


def _serialize_xml(root: ET.Element) -> bytes:
    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _add_manifest_entry(root: ET.Element, full_path: str, media_type: str) -> None:
    entry = ET.SubElement(root, f"{{{MANIFEST_NS}}}file-entry")
    entry.set(f"{{{MANIFEST_NS}}}full-path", full_path)
    entry.set(f"{{{MANIFEST_NS}}}media-type", media_type)


def _validate_replacement_paths(replacements: ReplacementMap) -> None:
    for name in replacements:
        _validate_member_path(name)


def _validate_member_path(name: str) -> None:
    path = PurePosixPath(name)
    if (
        not name
        or name.startswith("/")
        or "\\" in name
        or "\x00" in name
        or ".." in path.parts
        or any(part in {"", "."} for part in path.parts)
    ):
        raise OdsToolError(f"ODS package contains unsafe member path: {name!r}")


def _is_signature_entry(name: str) -> bool:
    lowered = name.casefold()
    return lowered.startswith("meta-inf/") and ("signature" in lowered or lowered.endswith(".p7s"))


def _compression_name(value: int) -> str:
    names = {
        zipfile.ZIP_STORED: "stored",
        zipfile.ZIP_DEFLATED: "deflated",
        zipfile.ZIP_BZIP2: "bzip2",
        zipfile.ZIP_LZMA: "lzma",
    }
    return names.get(value, f"unknown-{value}")


def _hash_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _diff_verified(
    before_path: Path,
    after_path: Path,
    before: PackageVerification,
    after: PackageVerification,
) -> PackageDiff:
    before_entries = {entry.path: entry.sha256 for entry in before.entries}
    after_entries = {entry.path: entry.sha256 for entry in after.entries}
    before_names = set(before_entries)
    after_names = set(after_entries)
    common = before_names & after_names
    return PackageDiff(
        before_sha256=_hash_path(before_path),
        after_sha256=_hash_path(after_path),
        added=sorted(after_names - before_names),
        removed=sorted(before_names - after_names),
        modified=sorted(name for name in common if before_entries[name] != after_entries[name]),
        unchanged=sum(before_entries[name] == after_entries[name] for name in common),
        signatures_before=before.signature_entries,
        signatures_after=after.signature_entries,
    )


def _content_changed(diff: PackageDiff) -> bool:
    signature_paths = set(diff.signatures_before) | set(diff.signatures_after)
    changed = set(diff.added) | set(diff.removed) | set(diff.modified)
    return bool(changed - signature_paths)


def _require_valid(verification: PackageVerification) -> None:
    if not verification.valid:
        raise OdsToolError("; ".join(verification.errors))


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1]


def _fsync_file(path: Path) -> None:
    with path.open("rb") as handle:
        os.fsync(handle.fileno())


def _fsync_directory(path: Path) -> None:
    descriptor = os.open(path, os.O_RDONLY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def _fsync_tree(root: Path) -> None:
    for path in sorted(root.rglob("*")):
        if path.is_file():
            _fsync_file(path)
    for path in sorted(
        (item for item in root.rglob("*") if item.is_dir()),
        key=lambda item: len(item.parts),
        reverse=True,
    ):
        _fsync_directory(path)
    _fsync_directory(root)


def _load_binding(path: Path) -> EventBindingSpec:
    try:
        payload = yaml.safe_load(path.read_text(encoding="utf-8"))
        return EventBindingSpec.model_validate(payload)
    except (OSError, UnicodeDecodeError, yaml.YAMLError, ValueError) as exc:
        raise OdsToolError(f"Invalid event binding YAML: {exc}") from exc


def _echo_model(model: BaseModel) -> None:
    click.echo(model.model_dump_json(indent=2))


def _mutation_options(function: Any) -> Any:
    for option in reversed(
        [
            click.option("--expect-sha256", type=str),
            click.option("--dry-run", is_flag=True),
        ]
    ):
        function = option(function)
    return function


@click.group()
@click.version_option(version="0.1.0")
def cli() -> None:
    """Safe, deterministic ODS package inspection and mutation."""
    try:
        require_application_container()
    except ContainerBoundaryError as exc:
        raise click.ClickException(str(exc)) from exc


@cli.command("list")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def list_command(ods_file: Path) -> None:
    """List verified ODS package members."""
    _run_cli(lambda: list_package(ods_file))


@cli.command("verify")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def verify_command(ods_file: Path) -> None:
    """Verify package and script/event invariants."""
    verification = verify_package(ods_file)
    _echo_model(verification)
    if not verification.valid:
        raise click.exceptions.Exit(1)


@cli.command("inspect-scripts")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def inspect_scripts_command(ods_file: Path) -> None:
    """Inspect embedded Python modules and exports."""
    _run_cli(lambda: inspect_scripts(ods_file))


@cli.command("upsert-script")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("module_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@_mutation_options
def upsert_script_command(
    ods_file: Path,
    module_file: Path,
    expect_sha256: str | None,
    dry_run: bool,
) -> None:
    """Upsert one Python module without deleting unrelated members."""
    _run_cli(
        lambda: upsert_scripts(
            ods_file,
            {module_file.name: module_file.read_text(encoding="utf-8")},
            expect_sha256=expect_sha256,
            dry_run=dry_run,
        )
    )


@cli.command("remove-script")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("module")
@_mutation_options
def remove_script_command(
    ods_file: Path,
    module: str,
    expect_sha256: str | None,
    dry_run: bool,
) -> None:
    """Remove one named Python module."""
    _run_cli(
        lambda: remove_scripts(
            ods_file,
            [module],
            expect_sha256=expect_sha256,
            dry_run=dry_run,
        )
    )


@cli.command("bind-event")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("binding_yaml", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@_mutation_options
def bind_event_command(
    ods_file: Path,
    binding_yaml: Path,
    expect_sha256: str | None,
    dry_run: bool,
) -> None:
    """Bind a YAML-declared event after resolving its target."""
    _run_cli(
        lambda: bind_event(
            ods_file,
            _load_binding(binding_yaml),
            expect_sha256=expect_sha256,
            dry_run=dry_run,
        )
    )


@cli.command("unbind-event")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("binding_id")
@_mutation_options
def unbind_event_command(
    ods_file: Path,
    binding_id: str,
    expect_sha256: str | None,
    dry_run: bool,
) -> None:
    """Remove one binding by its stable ID."""
    _run_cli(
        lambda: unbind_event(
            ods_file,
            binding_id,
            expect_sha256=expect_sha256,
            dry_run=dry_run,
        )
    )


@cli.command("diff")
@click.argument("before", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.argument("after", type=click.Path(exists=True, dir_okay=False, path_type=Path))
def diff_command(before: Path, after: Path) -> None:
    """Compare two verified ODS packages."""
    _run_cli(lambda: diff_packages(before, after))


@cli.command("snapshot")
@click.argument("ods_file", type=click.Path(exists=True, dir_okay=False, path_type=Path))
@click.option("--output", required=True, type=click.Path(file_okay=False, path_type=Path))
def snapshot_command(ods_file: Path, output: Path) -> None:
    """Create a transactional raw package snapshot."""
    try:
        destination = snapshot_package(ods_file, output)
    except OdsToolError as exc:
        raise click.ClickException(str(exc)) from exc
    click.echo(json.dumps({"snapshot": str(destination)}, indent=2))


def _run_cli(operation: Callable[[], BaseModel]) -> None:
    try:
        _echo_model(operation())
    except (OdsToolError, OSError, UnicodeDecodeError) as exc:
        raise click.ClickException(str(exc)) from exc


if __name__ == "__main__":
    cli()
