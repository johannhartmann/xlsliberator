"""Transactional evidence-bundle creation."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from datetime import UTC, datetime
from pathlib import Path
from uuid import uuid4

from pydantic import BaseModel

from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    EvidenceBundleManifest,
    RuntimeTrace,
    Scenario,
    TraceDiff,
)
from xlsliberator.validation_models import InventoryDiff, WorkbookArtifactIR


class EvidenceBundleWriter:
    """Write versioned scenario evidence with content hashes and relative references."""

    def __init__(self, directory: Path) -> None:
        self.directory = directory

    def write(
        self,
        *,
        source_workbook: Path,
        output: Path | None,
        environment: EnvironmentManifest,
        scenario: Scenario,
        source_trace: RuntimeTrace | None,
        target_traces: dict[str, RuntimeTrace],
        diffs: list[TraceDiff],
        source_inventory: WorkbookArtifactIR | None = None,
        target_inventories: dict[str, WorkbookArtifactIR] | None = None,
        inventory_diffs: list[InventoryDiff] | None = None,
    ) -> EvidenceBundleManifest:
        self.directory.mkdir(parents=True, exist_ok=True)
        environment_path = self._write_model("environment.json", environment)
        scenario_path = self._write_model("scenario.json", scenario)
        source_trace_path = (
            self._write_model("source-trace.json", source_trace) if source_trace else None
        )
        source_inventory_path = (
            self._write_model("source-inventory.json", source_inventory)
            if source_inventory
            else None
        )
        inventory_paths = {
            name: self._write_model(f"target-{name}-inventory.json", inventory)
            for name, inventory in sorted((target_inventories or {}).items())
        }
        inventory_diff_paths = [
            self._write_model(f"inventory-diff-{index}.json", difference)
            for index, difference in enumerate(inventory_diffs or [], start=1)
        ]
        target_paths = {
            name: self._write_model(f"target-{name}-trace.json", trace)
            for name, trace in sorted(target_traces.items())
        }
        diff_paths = [
            self._write_model(f"trace-diff-{index}.json", diff)
            for index, diff in enumerate(diffs, start=1)
        ]
        identities = {
            **({"source": source_trace.runtime_identity} if source_trace else {}),
            **{name: trace.runtime_identity for name, trace in target_traces.items()},
        }
        manifest = EvidenceBundleManifest(
            bundle_id=uuid4().hex,
            created_at=datetime.now(UTC),
            source_workbook_hash=_hash_file(source_workbook),
            output_hash=_hash_file(output) if output else None,
            environment_manifest=environment_path,
            scenario_definition=scenario_path,
            source_trace=source_trace_path,
            source_inventory=source_inventory_path,
            target_inventories=inventory_paths,
            inventory_diffs=inventory_diff_paths,
            target_traces=target_paths,
            trace_diffs=diff_paths,
            runtime_identities=identities,
            schema_versions={
                "environment": environment.schema_version,
                "scenario": scenario.schema_version,
                "runtime_trace": "1.0.0",
                "trace_diff": "1.0.0",
                "workbook_inventory": "3.0.0",
                "inventory_diff": "1.0.0",
            },
            granted_capabilities=sorted(environment.all_granted_capabilities),
        )
        self._write_model("manifest.json", manifest)
        return manifest

    def _write_model(self, name: str, model: BaseModel) -> str:
        path = self.directory / name
        payload = model.model_dump_json(indent=2)
        descriptor, temporary = tempfile.mkstemp(prefix=f".{name}.", dir=self.directory)
        try:
            with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
                handle.write(payload)
                handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, path)
        except Exception:
            Path(temporary).unlink(missing_ok=True)
            raise
        return name


def inspect_evidence_bundle(directory: Path) -> EvidenceBundleManifest:
    """Validate and return one evidence manifest without executing a runtime."""
    payload = json.loads((directory / "manifest.json").read_text(encoding="utf-8"))
    manifest = EvidenceBundleManifest.model_validate(payload)
    for reference in (
        manifest.environment_manifest,
        manifest.scenario_definition,
        *manifest.target_traces.values(),
        *manifest.trace_diffs,
        *manifest.target_inventories.values(),
        *manifest.inventory_diffs,
    ):
        if not (directory / reference).is_file():
            raise ValueError(f"evidence reference is missing: {reference}")
    if manifest.source_trace and not (directory / manifest.source_trace).is_file():
        raise ValueError(f"evidence reference is missing: {manifest.source_trace}")
    if manifest.source_inventory and not (directory / manifest.source_inventory).is_file():
        raise ValueError(f"evidence reference is missing: {manifest.source_inventory}")
    return manifest


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()
