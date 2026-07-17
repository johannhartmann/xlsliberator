#!/usr/bin/env python3
"""Reproducibly fetch, patch, build, and test the pinned LibreOffice source."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
WORK_ROOT = Path(os.environ.get("XLSLIBERATOR_OFFICE_WORK_ROOT", "/office-work"))
ARTIFACT_ROOT = Path(os.environ.get("XLSLIBERATOR_OFFICE_ARTIFACT_ROOT", "/office-artifacts"))
SOURCE_ROOT = WORK_ROOT / "sources" / "libreoffice"
WORKTREE_ROOT = WORK_ROOT / "worktrees"
TARBALL_ROOT = WORK_ROOT / "tarballs"


def _require_build_container() -> None:
    if (
        os.environ.get("XLSLIBERATOR_OFFICE_BUILD_CONTAINER") != "1"
        or not Path("/.dockerenv").is_file()
    ):
        raise RuntimeError(
            "tools/office.py may run only through the pinned Docker source-build service; "
            "use ./tools/office"
        )


def _run(
    command: list[str],
    *,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
    capture: bool = False,
) -> subprocess.CompletedProcess[str]:
    print("+", " ".join(command), flush=True)
    return subprocess.run(
        command,
        cwd=cwd or ROOT,
        env=env,
        check=True,
        text=True,
        capture_output=capture,
    )


def _load_manifest() -> dict[str, Any]:
    path = ROOT / "office" / "libreoffice" / "manifest.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema_version") != "1.0.0" or data.get("office_id") != "libreoffice":
        raise RuntimeError("unsupported office-source manifest")
    _verify_patch_series(data)
    return data


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _verify_patch_series(manifest: dict[str, Any]) -> None:
    series = manifest["patch_series"]
    series_path = ROOT / series["series_file"]
    names = [line.strip() for line in series_path.read_text().splitlines() if line.strip()]
    patches = series["patches"]
    expected = [Path(item["path"]).name for item in patches]
    if names != expected:
        raise RuntimeError("patch series and manifest differ")
    for item in patches:
        path = ROOT / item["path"]
        if _sha256(path) != item["sha256"]:
            raise RuntimeError(f"patch checksum mismatch: {path}")


def _safe_archive_members(archive: Path) -> None:
    listing = _run(["tar", "-tJf", str(archive)], capture=True).stdout.splitlines()
    if not listing:
        raise RuntimeError("source archive is empty")
    for member in listing:
        path = Path(member)
        if path.is_absolute() or ".." in path.parts:
            raise RuntimeError(f"unsafe source archive member: {member}")


def fetch(_args: argparse.Namespace) -> None:
    manifest = _load_manifest()
    archive_info = manifest["upstream"]["source_archive"]
    WORK_ROOT.mkdir(parents=True, exist_ok=True)
    TARBALL_ROOT.mkdir(parents=True, exist_ok=True)
    archive = WORK_ROOT / Path(archive_info["url"]).name
    if not archive.is_file() or _sha256(archive) != archive_info["sha256"]:
        temporary = archive.with_suffix(f"{archive.suffix}.part")
        temporary.unlink(missing_ok=True)
        _run(
            [
                "curl",
                "--fail",
                "--location",
                "--retry",
                "3",
                "--output",
                str(temporary),
                archive_info["url"],
            ]
        )
        if _sha256(temporary) != archive_info["sha256"]:
            raise RuntimeError("LibreOffice source archive checksum mismatch")
        temporary.replace(archive)
    _safe_archive_members(archive)

    source_ready = False
    try:
        _validate_source(manifest)
        marker = json.loads((SOURCE_ROOT / ".xlsliberator-source.json").read_text())
        source_ready = (
            marker.get("archive_sha256") == archive_info["sha256"]
            and marker.get("index_policy") == "include-release-generated-files-v1"
        )
    except RuntimeError:
        pass
    if not source_ready:
        _initialize_source_checkout(archive, archive_info, manifest)

    download_tree = _fresh_worktree("download-input", manifest, patched=False)
    try:
        _configure(download_tree, manifest)
        # Release source archives do not ship the development checkout's
        # top-level ./download helper.  The generated Makefile exposes the
        # supported tarball-aware entry point and also fetches/unpacks any
        # source modules required by this exact configuration.
        _run(["make", "fetch"], cwd=download_tree)
    finally:
        _remove_worktree(download_tree)
    print(json.dumps({"source": str(SOURCE_ROOT), "archive_sha256": _sha256(archive)}))


def _validate_source(manifest: dict[str, Any]) -> None:
    marker_path = SOURCE_ROOT / ".xlsliberator-source.json"
    if not marker_path.is_file():
        raise RuntimeError("source is not fetched; run ./tools/office fetch libreoffice")
    marker = json.loads(marker_path.read_text())
    if marker.get("upstream_commit") != manifest["upstream"]["commit"]:
        raise RuntimeError("fetched source commit identity differs from manifest")
    status = _run(["git", "status", "--porcelain"], cwd=SOURCE_ROOT, capture=True).stdout
    if status:
        raise RuntimeError("baseline source checkout is dirty")


def _initialize_source_checkout(
    archive: Path, archive_info: dict[str, Any], manifest: dict[str, Any]
) -> None:
    """Create the deterministic local baseline from the verified release archive."""
    with tempfile.TemporaryDirectory(dir=WORK_ROOT, prefix="source-extract-") as temp_dir:
        extracted = Path(temp_dir) / "source"
        extracted.mkdir()
        _run(
            [
                "tar",
                "-xJf",
                str(archive),
                "--strip-components=1",
                "--no-same-owner",
                "-C",
                str(extracted),
            ]
        )
        if not (extracted / "COPYING").is_file() or not (extracted / "autogen.sh").is_file():
            raise RuntimeError("source archive lacks expected LibreOffice files")
        marker = {
            "schema_version": "1.0.0",
            "archive_sha256": archive_info["sha256"],
            "upstream_commit": manifest["upstream"]["commit"],
            "upstream_tag": manifest["upstream"]["tag"],
            "index_policy": "include-release-generated-files-v1",
        }
        (extracted / ".xlsliberator-source.json").write_text(
            json.dumps(marker, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
        _run(["git", "init", "--initial-branch=baseline"], cwd=extracted)
        _run(["git", "config", "user.name", "XLSLiberator Office Build"], cwd=extracted)
        _run(["git", "config", "user.email", "office-build@invalid.local"], cwd=extracted)
        # Release tarballs contain ignored-but-required bootstrap files such as
        # configure and config.guess. The local baseline must preserve the full
        # verified archive, not Git's upstream development-checkout ignore view.
        _run(["git", "add", "--all", "--force"], cwd=extracted)
        fixed_env = dict(os.environ)
        fixed_env.update(
            {
                "GIT_AUTHOR_DATE": "2026-06-10T12:00:00Z",
                "GIT_COMMITTER_DATE": "2026-06-10T12:00:00Z",
            }
        )
        _run(
            [
                "git",
                "commit",
                "--quiet",
                "-m",
                f"LibreOffice {manifest['baseline']['full_build']} baseline",
            ],
            cwd=extracted,
            env=fixed_env,
        )
        _run(["git", "tag", "xlsliberator-baseline"], cwd=extracted)
        if SOURCE_ROOT.exists():
            shutil.rmtree(SOURCE_ROOT)
        SOURCE_ROOT.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(extracted), SOURCE_ROOT)


def _remove_worktree(path: Path) -> None:
    if path.exists():
        _run(["git", "worktree", "remove", "--force", str(path)], cwd=SOURCE_ROOT)


def _fresh_worktree(name: str, manifest: dict[str, Any], *, patched: bool) -> Path:
    _validate_source(manifest)
    path = WORKTREE_ROOT / name
    _remove_worktree(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    _run(
        ["git", "worktree", "add", "--quiet", "--detach", str(path), "xlsliberator-baseline"],
        cwd=SOURCE_ROOT,
    )
    if patched:
        for item in manifest["patch_series"]["patches"]:
            _run(["git", "apply", "--index", str(ROOT / item["path"])], cwd=path)
    return path


def _configure(source: Path, manifest: dict[str, Any], *, parallelism: int = 1) -> None:
    if parallelism < 1:
        raise ValueError("office build parallelism must be at least one")
    command = [
        "perl",
        "./autogen.sh",
        f"--with-external-tar={TARBALL_ROOT}",
        f"--with-parallelism={parallelism}",
        *manifest["build"]["options"],
    ]
    _run(command, cwd=source)


def _compiler_identity() -> dict[str, str]:
    identities: dict[str, str] = {}
    for executable in ("gcc", "g++", "ld", "make"):
        result = _run([executable, "--version"], capture=True)
        identities[executable] = result.stdout.splitlines()[0]
    return identities


def _hash_runtime_binaries(program_dir: Path) -> dict[str, str]:
    patterns = (
        "soffice",
        "soffice.bin",
        "python",
        "python.bin",
        "pyuno*.so",
        "libsclo.so",
        "libscfiltlo.so",
    )
    result: dict[str, str] = {}
    for pattern in patterns:
        for path in sorted(program_dir.glob(pattern)):
            if path.is_file():
                result[str(path.relative_to(program_dir))] = _sha256(path)
    required = {"soffice", "soffice.bin"}
    if not required.issubset(result):
        raise RuntimeError("source build did not produce the required office binaries")
    return result


def build(args: argparse.Namespace) -> None:
    manifest = _load_manifest()
    patched = args.variant == "patched"
    variant = manifest["patch_series"]["variant"] if patched else "stock-source"
    source = _fresh_worktree(f"build-{variant}", manifest, patched=patched)
    _configure(source, manifest, parallelism=args.jobs)
    env = dict(os.environ)
    env["CCACHE_BASEDIR"] = str(source)
    _run(["make", f"-j{args.jobs}"], cwd=source, env=env)
    program_dir = source / "instdir" / "program"
    binary_hashes = _hash_runtime_binaries(program_dir)
    artifact_dir = ARTIFACT_ROOT / "office-build" / "libreoffice" / variant
    artifact_dir.mkdir(parents=True, exist_ok=True)
    package_manifest = _run(
        ["dpkg-query", "-W", "-f=${Package}\t${Version}\n"], capture=True
    ).stdout.splitlines()
    identity = {
        "schema_version": "1.0.0",
        "status": "built-not-yet-conformance-tested",
        "office_id": "libreoffice",
        "libreoffice_build": manifest["baseline"]["full_build"],
        "runtime_variant": variant,
        "source_tag": manifest["upstream"]["tag"],
        "source_commit": manifest["upstream"]["commit"],
        "source_archive_sha256": manifest["upstream"]["source_archive"]["sha256"],
        "patches": manifest["patch_series"]["patches"] if patched else [],
        "build_base_image": manifest["build"]["base_image"],
        "debian_snapshot": manifest["build"]["debian_snapshot"],
        "architecture": platform.machine(),
        "compiler": _compiler_identity(),
        "build_options": manifest["build"]["options"],
        "package_manifest": package_manifest,
        "binary_sha256": binary_hashes,
        "runtime_image_reference": (
            manifest["result"]["runtime_image"]
            if patched
            else "xlsliberator-libreoffice:26.2.4.2-stock-source"
        ),
        "runtime_image_digest": None,
        "python_version": None,
        "pyuno_identity": None,
        "built_at": datetime.now(UTC).isoformat(),
    }
    (artifact_dir / "identity.json").write_text(
        json.dumps(identity, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    archive = artifact_dir / "instdir.tar.gz"
    _run(
        [
            "tar",
            "--sort=name",
            "--mtime=@0",
            "--owner=0",
            "--group=0",
            "--numeric-owner",
            "--use-compress-program=gzip -n",
            "-cf",
            str(archive),
            "-C",
            str(source),
            "instdir",
        ]
    )
    print(json.dumps({"identity": str(artifact_dir / "identity.json"), "artifact": str(archive)}))


def test(args: argparse.Namespace) -> None:
    manifest = _load_manifest()
    variant = manifest["patch_series"]["variant"] if args.variant == "patched" else "stock-source"
    source = WORKTREE_ROOT / f"build-{variant}"
    program = source / "instdir" / "program"
    python = program / "python"
    if not python.is_file():
        raise RuntimeError("built LibreOffice Python is absent; run build first")
    fixture = json.loads((ROOT / manifest["conformance"]["fixture"]).read_text())
    env = dict(os.environ)
    env.update(
        {
            "PYTHONPATH": f"{program}:{ROOT / 'src'}",
            "XLSLIBERATOR_OFFICE_CONTAINER": "1",
            "XLSLIBERATOR_SOURCE_BUILD_CONTAINER": "1",
            "XLSLIBERATOR_OFFICE_PYTHON_PREFIX": str(program),
            "XLSLIBERATOR_OFFICE_EXECUTABLE": str(program / "soffice"),
        }
    )
    request = json.dumps({"op": "evaluate_formula", "formula": fixture["formula"]})
    result = subprocess.run(
        [str(python), "-m", "xlsliberator.lo_worker"],
        input=request,
        cwd=source,
        env=env,
        check=True,
        text=True,
        capture_output=True,
    )
    response = json.loads(result.stdout)
    observed = response.get("data") if response.get("success") else None
    expected = fixture["expected"]
    passed = bool(
        observed
        and observed.get("error_code") == expected["error_code"]
        and observed.get("string") == expected["string"]
    )
    if args.variant == "patched" and not passed:
        raise RuntimeError("patched source runtime failed the upstream conformance case")
    if args.variant == "stock" and passed:
        raise RuntimeError(
            "stock source unexpectedly passes; regression case is no longer discriminating"
        )
    evidence = {
        "schema_version": "1.0.0",
        "case_id": fixture["case_id"],
        "variant": variant,
        "expected": expected,
        "observed": observed,
        "disposition": "passed" if passed else "failed",
        "worker_stderr": result.stderr[-16000:],
    }
    evidence_dir = ARTIFACT_ROOT / "office-build" / "libreoffice" / variant
    evidence_dir.mkdir(parents=True, exist_ok=True)
    (evidence_dir / "conformance.json").write_text(
        json.dumps(evidence, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    identity_path = evidence_dir / "identity.json"
    if identity_path.is_file():
        identity = json.loads(identity_path.read_text(encoding="utf-8"))
        identity["status"] = (
            "patched-conformance-passed" if passed else "stock-regression-confirmed"
        )
        identity["conformance_evidence"] = str(evidence_dir / "conformance.json")
        identity["conformance_tested_at"] = datetime.now(UTC).isoformat()
        identity_path.write_text(
            json.dumps(identity, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
    print(json.dumps(evidence, sort_keys=True))


def record_runtime(args: argparse.Namespace) -> None:
    """Merge a Docker-orchestrated immutable image probe into build evidence."""
    manifest = _load_manifest()
    variant = manifest["patch_series"]["variant"] if args.variant == "patched" else "stock-source"
    evidence_dir = ARTIFACT_ROOT / "office-build" / "libreoffice" / variant
    identity_path = evidence_dir / "identity.json"
    if not identity_path.is_file():
        raise RuntimeError("build identity is absent; run build first")
    if not args.image_id.startswith("sha256:") or len(args.image_id) != 71:
        raise RuntimeError("runtime image did not resolve to an immutable sha256 identity")
    response = json.loads(args.probe_json)
    if not response.get("success"):
        raise RuntimeError("source runtime probe failed")
    probe = dict(response.get("data") or {})
    if probe.get("libreoffice_build") != manifest["baseline"]["full_build"]:
        raise RuntimeError("source runtime probe returned the wrong LibreOffice build")
    if probe.get("runtime_variant") != variant:
        raise RuntimeError("source runtime probe returned the wrong runtime variant")
    expected_patch = (
        manifest["patch_series"]["patches"][0]["sha256"] if args.variant == "patched" else "none"
    )
    if probe.get("source_commit") != manifest["upstream"]["commit"]:
        raise RuntimeError("source runtime probe returned the wrong source commit")
    if probe.get("patch_set_sha256") != expected_patch:
        raise RuntimeError("source runtime probe returned the wrong patch identity")
    identity = json.loads(identity_path.read_text(encoding="utf-8"))
    identity.update(
        {
            "status": "runtime-probed-not-yet-conformance-tested",
            "runtime_image_digest": args.image_id,
            "runtime_architecture": probe.get("architecture"),
            "python_version": probe.get("python_version"),
            "pyuno_identity": {
                "uno_module": probe.get("uno_module"),
                "uno_module_sha256": probe.get("uno_module_sha256"),
                "pyuno_native_module": probe.get("pyuno_native_module"),
                "pyuno_native_sha256": probe.get("pyuno_native_sha256"),
            },
            "runtime_binary_sha256": {
                "office": probe.get("office_sha256"),
                "worker_wrapper": probe.get("worker_wrapper_sha256"),
            },
            "runtime_package_manifest": probe.get("installed_package_manifest"),
            "runtime_probed_at": datetime.now(UTC).isoformat(),
        }
    )
    identity_path.write_text(
        json.dumps(identity, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(json.dumps({"identity": str(identity_path), "image_id": args.image_id}))


def worktree(args: argparse.Namespace) -> None:
    manifest = _load_manifest()
    safe_name = args.name.replace("/", "-").replace("..", "-")
    if not safe_name or safe_name.startswith("-"):
        raise RuntimeError("invalid worktree name")
    path = _fresh_worktree(f"agent-{safe_name}", manifest, patched=args.with_patches)
    _run(["git", "switch", "-c", f"agent/{safe_name}"], cwd=path)
    print(path)


def main() -> int:
    _require_build_container()
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="action", required=True)
    for action, handler in (
        ("fetch", fetch),
        ("build", build),
        ("test", test),
        ("worktree", worktree),
        ("record-runtime", record_runtime),
    ):
        command = subparsers.add_parser(action)
        command.add_argument("office", choices=["libreoffice"])
        if action in {"build", "test"}:
            command.add_argument("--variant", choices=["stock", "patched"], default="patched")
        if action == "build":
            command.add_argument("--jobs", type=int, default=2)
        if action == "worktree":
            command.add_argument("--name", required=True)
            command.add_argument("--with-patches", action="store_true")
        if action == "record-runtime":
            command.add_argument("--variant", choices=["stock", "patched"], required=True)
            command.add_argument("--image-id", required=True)
            command.add_argument("--probe-json", required=True)
        command.set_defaults(handler=handler)
    args = parser.parse_args()
    args.handler(args)
    return 0


if __name__ == "__main__":
    sys.exit(main())
