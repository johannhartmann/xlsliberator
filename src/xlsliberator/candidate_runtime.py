"""Generic, content-bound migration-candidate execution inside Docker.

Candidate bundles are generated migration artifacts.  They contain direct
target-native Python/UNO code for one workbook, but the XLSLiberator runtime
knows only the versioned bundle and entrypoint contract defined here.
"""

from __future__ import annotations

import hashlib
import importlib
import json
import re
import stat
import sys
import tempfile
from collections.abc import Callable, Iterator, Mapping
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from types import ModuleType
from typing import Any, Final
from zipfile import ZIP_STORED, BadZipFile, ZipFile, ZipInfo

from xlsliberator.office_target import LIBREOFFICE_VERSION

_SCHEMA_VERSION: Final = "1.0.0"
_MAX_FILES: Final = 100
_MAX_FILE_BYTES: Final = 4 * 1024 * 1024
_MAX_TOTAL_BYTES: Final = 16 * 1024 * 1024
_SHA256 = re.compile(r"^[0-9a-f]{64}$")
_CANDIDATE_ID = re.compile(r"^[a-z0-9][a-z0-9._-]{1,127}$")
_ENTRYPOINT = re.compile(
    r"^(?P<module>candidate_[a-zA-Z0-9_]*(?:\.[a-zA-Z_][a-zA-Z0-9_]*)*):"
    r"(?P<callable>[a-zA-Z_][a-zA-Z0-9_]*)$"
)


class CandidateBundleError(RuntimeError):
    """Raised when a generated migration candidate violates its contract."""


@dataclass(frozen=True, slots=True)
class CandidateManifest:
    """Validated identity and entrypoints for one generated candidate."""

    candidate_id: str
    source_sha256: str
    target_build: str
    build_entrypoint: str
    controller_entrypoint: str | None
    files: Mapping[str, str]
    capabilities: tuple[str, ...]


@dataclass(frozen=True, slots=True)
class LoadedEntrypoint:
    """One imported callable retained for the lifetime of its extraction root."""

    manifest: CandidateManifest
    bundle_sha256: str
    callback: Callable[..., Any]


def build_application_target(request: dict[str, Any]) -> dict[str, Any]:
    """Build one ODS through a generated candidate's declared builder."""
    source = Path(str(request["input_path"])).resolve()
    candidate_path = Path(str(request["candidate_path"])).resolve()
    output = Path(str(request["output_path"])).resolve()
    if output.suffix.lower() != ".ods":
        raise ValueError("migration candidate output must use the ODS format")

    source_before = _sha256_file(source)
    output.unlink(missing_ok=True)
    try:
        with load_candidate_entrypoint(candidate_path, "build") as loaded:
            if source_before != loaded.manifest.source_sha256:
                raise CandidateBundleError(
                    "candidate source identity does not match the supplied workbook"
                )
            result = loaded.callback(
                {
                    **request,
                    "input_path": str(source),
                    "candidate_path": str(candidate_path),
                    "output_path": str(output),
                    "candidate_id": loaded.manifest.candidate_id,
                    "candidate_bundle_sha256": loaded.bundle_sha256,
                }
            )
            if not isinstance(result, dict):
                raise CandidateBundleError("candidate builder returned a non-object result")
            if not output.is_file():
                raise CandidateBundleError("candidate builder did not produce its declared ODS")
            if _sha256_file(source) != source_before:
                raise CandidateBundleError(
                    "candidate builder mutated the immutable source workbook"
                )
            return {
                **result,
                "status": "passed",
                "candidate_id": loaded.manifest.candidate_id,
                "candidate_bundle_sha256": loaded.bundle_sha256,
                "source_sha256": source_before,
                "target_sha256": _sha256_file(output),
                "target_build": loaded.manifest.target_build,
                "target_format": "ods",
                "declared_capabilities": list(loaded.manifest.capabilities),
                "candidate_files": dict(sorted(loaded.manifest.files.items())),
            }
    except Exception:
        output.unlink(missing_ok=True)
        raise


def package_candidate_directory(source_directory: Path, output_path: Path) -> dict[str, Any]:
    """Create a deterministic bundle from a generated candidate directory."""
    source = source_directory.resolve()
    manifest_path = source / "manifest.json"
    if not source.is_dir() or not manifest_path.is_file():
        raise CandidateBundleError("candidate directory must contain manifest.json")
    try:
        raw_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise CandidateBundleError("candidate manifest is not valid JSON") from exc
    manifest = _manifest_from_payload(raw_manifest)
    for name, expected_sha256 in manifest.files.items():
        path = source.joinpath(*PurePosixPath(name).parts)
        if not path.is_file() or _sha256_file(path) != expected_sha256:
            raise CandidateBundleError(f"candidate file digest mismatch: {name}")

    output = output_path.resolve()
    if output.suffix.lower() != ".zip":
        raise CandidateBundleError("candidate bundle output must use the ZIP format")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.unlink(missing_ok=True)
    try:
        with ZipFile(output, "w", compression=ZIP_STORED) as archive:
            _write_deterministic_member(archive, "manifest.json", manifest_path.read_bytes())
            for name in sorted(manifest.files):
                path = source.joinpath(*PurePosixPath(name).parts)
                _write_deterministic_member(archive, name, path.read_bytes())
        with tempfile.TemporaryDirectory(prefix="xlsliberator-candidate-check-") as temporary:
            checked = _extract_and_validate(output, Path(temporary))
    except Exception:
        output.unlink(missing_ok=True)
        raise
    return {
        "status": "passed",
        "candidate_id": checked.candidate_id,
        "source_sha256": checked.source_sha256,
        "target_build": checked.target_build,
        "candidate_bundle_sha256": _sha256_file(output),
        "candidate_files": dict(sorted(checked.files.items())),
        "output_bytes": output.stat().st_size,
    }


@contextmanager
def load_candidate_entrypoint(
    bundle_path: Path,
    role: str,
) -> Iterator[LoadedEntrypoint]:
    """Validate, extract, and import exactly one declared candidate entrypoint."""
    bundle = bundle_path.resolve()
    bundle_sha256 = _sha256_file(bundle)
    with tempfile.TemporaryDirectory(prefix="xlsliberator-candidate-") as temporary:
        root = Path(temporary)
        manifest = _extract_and_validate(bundle, root)
        entrypoint = {
            "build": manifest.build_entrypoint,
            "controller": manifest.controller_entrypoint,
        }.get(role)
        if not entrypoint:
            raise CandidateBundleError(f"candidate does not declare a {role!r} entrypoint")
        callback, package_name = _import_entrypoint(root, entrypoint)
        sys.path.insert(0, str(root))
        try:
            yield LoadedEntrypoint(
                manifest=manifest,
                bundle_sha256=bundle_sha256,
                callback=callback,
            )
        finally:
            with suppress(ValueError):
                sys.path.remove(str(root))
            for name in tuple(sys.modules):
                if name == package_name or name.startswith(f"{package_name}."):
                    sys.modules.pop(name, None)


def _extract_and_validate(bundle: Path, root: Path) -> CandidateManifest:
    if not bundle.is_file() or not bundle.name.lower().endswith(".zip"):
        raise CandidateBundleError("candidate bundle must be an existing ZIP file")
    try:
        with ZipFile(bundle) as archive:
            infos = archive.infolist()
            if not 2 <= len(infos) <= _MAX_FILES:
                raise CandidateBundleError("candidate bundle has an invalid file count")
            names: set[str] = set()
            total = 0
            for info in infos:
                name = _safe_member_name(info.filename)
                mode = info.external_attr >> 16
                if info.is_dir() or stat.S_ISLNK(mode) or info.flag_bits & 0x1 or name in names:
                    raise CandidateBundleError("candidate bundle contains an unsafe member")
                if info.file_size > _MAX_FILE_BYTES:
                    raise CandidateBundleError(f"candidate member is oversized: {name}")
                total += info.file_size
                if total > _MAX_TOTAL_BYTES:
                    raise CandidateBundleError("candidate bundle exceeds its expanded size limit")
                names.add(name)
            if "manifest.json" not in names:
                raise CandidateBundleError("candidate bundle has no manifest.json")
            manifest_payload = archive.read("manifest.json")
            try:
                raw_manifest = json.loads(manifest_payload)
            except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                raise CandidateBundleError("candidate manifest is not valid JSON") from exc
            manifest = _manifest_from_payload(raw_manifest)
            expected = {"manifest.json", *manifest.files}
            if names != expected:
                raise CandidateBundleError(
                    "candidate manifest file inventory does not match the archive"
                )
            for name, expected_sha256 in manifest.files.items():
                payload = archive.read(name)
                if hashlib.sha256(payload).hexdigest() != expected_sha256:
                    raise CandidateBundleError(f"candidate file digest mismatch: {name}")
                destination = root.joinpath(*PurePosixPath(name).parts)
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_bytes(payload)
    except BadZipFile as exc:
        raise CandidateBundleError("candidate bundle is not a valid ZIP file") from exc
    return manifest


def _manifest_from_payload(raw: object) -> CandidateManifest:
    if not isinstance(raw, dict):
        raise CandidateBundleError("candidate manifest must be an object")
    expected_keys = {
        "schema_version",
        "candidate_id",
        "source_sha256",
        "target_build",
        "entrypoints",
        "files",
        "capabilities",
    }
    if set(raw) != expected_keys:
        raise CandidateBundleError("candidate manifest has missing or unknown fields")
    if raw["schema_version"] != _SCHEMA_VERSION:
        raise CandidateBundleError("candidate manifest schema version is unsupported")

    candidate_id = raw["candidate_id"]
    source_sha256 = raw["source_sha256"]
    target_build = raw["target_build"]
    if not all(isinstance(value, str) for value in (candidate_id, source_sha256, target_build)):
        raise CandidateBundleError("candidate identity fields must be strings")
    if _CANDIDATE_ID.fullmatch(candidate_id) is None:
        raise CandidateBundleError("candidate_id is malformed")
    if _SHA256.fullmatch(source_sha256) is None:
        raise CandidateBundleError("candidate source digest is malformed")
    if target_build != LIBREOFFICE_VERSION:
        raise CandidateBundleError(
            f"candidate targets LibreOffice {target_build}, expected {LIBREOFFICE_VERSION}"
        )

    entrypoints = raw["entrypoints"]
    if not isinstance(entrypoints, dict) or set(entrypoints) != {"build", "controller"}:
        raise CandidateBundleError("candidate entrypoints must declare build and controller")
    build_entrypoint = _validated_entrypoint(entrypoints["build"], required=True)
    controller_entrypoint = _validated_entrypoint(entrypoints["controller"], required=False)
    if build_entrypoint is None:
        raise CandidateBundleError("candidate build entrypoint is required")

    raw_files = raw["files"]
    if not isinstance(raw_files, dict) or not raw_files:
        raise CandidateBundleError("candidate files must be a non-empty object")
    files: dict[str, str] = {}
    for raw_name, raw_digest in raw_files.items():
        if not isinstance(raw_name, str):
            raise CandidateBundleError("candidate file names must be strings")
        name = _safe_member_name(raw_name)
        if not isinstance(raw_digest, str):
            raise CandidateBundleError(f"candidate file digest is malformed: {name}")
        digest = raw_digest
        if _SHA256.fullmatch(digest) is None:
            raise CandidateBundleError(f"candidate file digest is malformed: {name}")
        files[name] = digest

    for entrypoint in (build_entrypoint, controller_entrypoint):
        if entrypoint is None:
            continue
        match = _ENTRYPOINT.fullmatch(entrypoint)
        if match is None:
            raise CandidateBundleError("candidate entrypoint is malformed")
        module_path = f"{match.group('module').replace('.', '/')}.py"
        package_init = f"{match.group('module').partition('.')[0]}/__init__.py"
        if module_path not in files or package_init not in files:
            raise CandidateBundleError(
                "candidate entrypoint module or package is absent from its inventory"
            )

    capabilities = raw["capabilities"]
    if (
        not isinstance(capabilities, list)
        or not all(isinstance(item, str) and 1 <= len(item) <= 100 for item in capabilities)
        or len(capabilities) != len(set(capabilities))
    ):
        raise CandidateBundleError("candidate capabilities must be unique non-empty strings")
    return CandidateManifest(
        candidate_id=candidate_id,
        source_sha256=source_sha256,
        target_build=target_build,
        build_entrypoint=build_entrypoint,
        controller_entrypoint=controller_entrypoint,
        files=files,
        capabilities=tuple(capabilities),
    )


def _validated_entrypoint(raw: object, *, required: bool) -> str | None:
    if raw is None and not required:
        return None
    if not isinstance(raw, str) or _ENTRYPOINT.fullmatch(raw) is None:
        raise CandidateBundleError("candidate entrypoint is malformed")
    return raw


def _import_entrypoint(root: Path, entrypoint: str) -> tuple[Callable[..., Any], str]:
    match = _ENTRYPOINT.fullmatch(entrypoint)
    if match is None:
        raise CandidateBundleError("candidate entrypoint is malformed")
    module_name = match.group("module")
    callable_name = match.group("callable")
    package_name = module_name.partition(".")[0]
    module_path = root.joinpath(*module_name.split(".")).with_suffix(".py")
    package_init = root / package_name / "__init__.py"
    if not module_path.is_file() or not package_init.is_file():
        raise CandidateBundleError("candidate entrypoint module or package is missing")

    preexisting = {
        name for name in sys.modules if name == package_name or name.startswith(f"{package_name}.")
    }
    if preexisting:
        raise CandidateBundleError("candidate package name collides with an imported module")
    sys.path.insert(0, str(root))
    try:
        module: ModuleType = importlib.import_module(module_name)
        callback = getattr(module, callable_name, None)
        if not callable(callback):
            raise CandidateBundleError("candidate entrypoint does not resolve to a callable")
        return callback, package_name
    except Exception:
        for name in tuple(sys.modules):
            if name == package_name or name.startswith(f"{package_name}."):
                sys.modules.pop(name, None)
        raise
    finally:
        sys.path.remove(str(root))


def _safe_member_name(raw_name: str) -> str:
    path = PurePosixPath(raw_name)
    if (
        not raw_name
        or path.is_absolute()
        or ".." in path.parts
        or "." in path.parts
        or "\\" in raw_name
        or raw_name != path.as_posix()
    ):
        raise CandidateBundleError("candidate bundle contains an unsafe path")
    return raw_name


def _write_deterministic_member(archive: ZipFile, name: str, payload: bytes) -> None:
    info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = ZIP_STORED
    info.external_attr = 0o100644 << 16
    archive.writestr(info, payload)


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


__all__ = [
    "CandidateBundleError",
    "CandidateManifest",
    "LoadedEntrypoint",
    "build_application_target",
    "load_candidate_entrypoint",
    "package_candidate_directory",
]
