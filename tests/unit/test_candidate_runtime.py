"""Security and lifecycle tests for generic migration-candidate bundles."""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsliberator.candidate_runtime import (
    CandidateBundleError,
    build_application_target,
    create_candidate_workspace,
    inspect_candidate_package,
    load_candidate_entrypoint,
    package_candidate_directory,
    package_candidate_workspace,
    validate_candidate_workspace,
)


def _candidate_directory(tmp_path: Path, source_sha256: str) -> Path:
    root = tmp_path / "candidate"
    package = root / "candidate_fixture_contract"
    package.mkdir(parents=True)
    files = {
        "candidate_fixture_contract/__init__.py": "",
        "candidate_fixture_contract/main.py": (
            "def build(request):\n"
            "    from candidate_fixture_contract.result import value\n"
            "    return {'value': value(), 'request': request}\n\n"
            "def controller(session, document, config):\n"
            "    return (session, document, config)\n"
        ),
        "candidate_fixture_contract/result.py": ("def value():\n    return 'loaded-lazily'\n"),
    }
    digests: dict[str, str] = {}
    for name, payload in files.items():
        path = root / name
        path.write_text(payload, encoding="utf-8")
        digests[name] = hashlib.sha256(payload.encode()).hexdigest()
    manifest = {
        "schema_version": "1.0.0",
        "candidate_id": "fixture-contract",
        "source_sha256": source_sha256,
        "target_build": "26.2.4.2",
        "entrypoints": {
            "build": "candidate_fixture_contract.main:build",
            "controller": "candidate_fixture_contract.main:controller",
        },
        "files": digests,
        "capabilities": [],
    }
    (root / "manifest.json").write_text(
        json.dumps(manifest, indent=2) + "\n",
        encoding="utf-8",
    )
    return root


def test_agent_workspace_tools_calculate_source_and_file_bindings(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsm"
    source.write_bytes(b"source-workbook")
    workspace = tmp_path / "authored"

    scaffold = create_candidate_workspace(source, workspace, candidate_id="payroll-migration")
    module_path = workspace / "candidate_payroll_migration/main.py"
    module_path.write_text(
        "def build(request):\n    return {'request': request}\n",
        encoding="utf-8",
    )
    manifest_path = workspace / "manifest.json"
    draft = json.loads(manifest_path.read_text(encoding="utf-8"))
    draft["source_sha256"] = "0" * 64
    draft["files"] = {"invented.py": "0" * 64}
    manifest_path.write_text(json.dumps(draft), encoding="utf-8")

    validated = validate_candidate_workspace(source, workspace)
    bundle = tmp_path / "candidate.zip"
    packaged = package_candidate_workspace(source, workspace, bundle)
    inspected = inspect_candidate_package(bundle)

    source_sha256 = hashlib.sha256(source.read_bytes()).hexdigest()
    module_sha256 = hashlib.sha256(module_path.read_bytes()).hexdigest()
    assert scaffold["source_sha256"] == source_sha256
    assert validated["source_sha256"] == source_sha256
    assert validated["candidate_files"]["candidate_payroll_migration/main.py"] == module_sha256
    assert packaged["source_sha256"] == source_sha256
    assert inspected["candidate_bundle_sha256"] == packaged["candidate_bundle_sha256"]
    assert "invented.py" not in inspected["candidate_files"]


def test_candidate_package_is_deterministic_and_keeps_lazy_imports_confined(
    tmp_path: Path,
) -> None:
    source_sha256 = hashlib.sha256(b"source").hexdigest()
    directory = _candidate_directory(tmp_path, source_sha256)
    first = tmp_path / "first.zip"
    second = tmp_path / "second.zip"

    first_result = package_candidate_directory(directory, first)
    second_result = package_candidate_directory(directory, second)

    assert first.read_bytes() == second.read_bytes()
    assert first_result["candidate_bundle_sha256"] == second_result["candidate_bundle_sha256"]
    with load_candidate_entrypoint(first, "build") as loaded:
        result = loaded.callback({"input_path": "opaque"})
        assert result["value"] == "loaded-lazily"
        assert "candidate_fixture_contract.result" in sys.modules
    assert not any(
        name == "candidate_fixture_contract" or name.startswith("candidate_fixture_contract.")
        for name in sys.modules
    )


def test_candidate_rejects_source_identity_mismatch_before_running_builder(
    tmp_path: Path,
) -> None:
    source = tmp_path / "source.xlsb"
    source.write_bytes(b"source")
    directory = _candidate_directory(tmp_path, hashlib.sha256(b"different").hexdigest())
    bundle = tmp_path / "candidate.zip"
    package_candidate_directory(directory, bundle)

    with pytest.raises(CandidateBundleError, match="source identity"):
        build_application_target(
            {
                "input_path": str(source),
                "candidate_path": str(bundle),
                "output_path": str(tmp_path / "target.ods"),
            }
        )


def test_candidate_rejects_unmanifested_and_traversal_members(tmp_path: Path) -> None:
    source_sha256 = hashlib.sha256(b"source").hexdigest()
    directory = _candidate_directory(tmp_path, source_sha256)
    bundle = tmp_path / "candidate.zip"
    package_candidate_directory(directory, bundle)
    with ZipFile(bundle, "a") as archive:
        archive.writestr("../escape.py", b"raise AssertionError")

    with pytest.raises(CandidateBundleError, match="unsafe path"):
        with load_candidate_entrypoint(bundle, "build"):
            pass


def test_candidate_package_rejects_changed_generated_file(tmp_path: Path) -> None:
    directory = _candidate_directory(tmp_path, hashlib.sha256(b"source").hexdigest())
    (directory / "candidate_fixture_contract/result.py").write_text(
        "def value():\n    return 'changed'\n",
        encoding="utf-8",
    )

    with pytest.raises(CandidateBundleError, match="digest mismatch"):
        package_candidate_directory(directory, tmp_path / "candidate.zip")


def test_candidate_manifest_rejects_coerced_identity_and_missing_entrypoint(
    tmp_path: Path,
) -> None:
    directory = _candidate_directory(tmp_path, hashlib.sha256(b"source").hexdigest())
    manifest_path = directory / "manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["candidate_id"] = 42
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    with pytest.raises(CandidateBundleError, match="identity fields must be strings"):
        package_candidate_directory(directory, tmp_path / "coerced.zip")

    manifest["candidate_id"] = "fixture-contract"
    manifest["entrypoints"]["build"] = "candidate_fixture_contract.missing:build"
    manifest_path.write_text(json.dumps(manifest), encoding="utf-8")

    with pytest.raises(CandidateBundleError, match="entrypoint module or package is absent"):
        package_candidate_directory(directory, tmp_path / "missing-entrypoint.zip")
