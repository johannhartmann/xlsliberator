"""Versioned LibreOffice source, patch-series, and runtime selection models."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)


class SourceArchive(StrictModel):
    url: str
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")


class UpstreamSource(StrictModel):
    repository: str
    fetch_mirror: str
    tag: str
    tag_object: str = Field(pattern=r"^[0-9a-f]{40}$")
    commit: str = Field(pattern=r"^[0-9a-f]{40}$")
    tree: str = Field(pattern=r"^[0-9a-f]{40}$")
    license: str
    license_file: str
    source_archive: SourceArchive


class StockBaseline(StrictModel):
    full_build: Literal["26.2.4.2"]
    runtime_variant: Literal["stock"]
    runtime_image: str
    lock_file: str


class BuildDefinition(StrictModel):
    dockerfile: str
    base_image: str
    debian_snapshot: str
    toolchain: str
    options: list[str]
    correctness_cache_policy: str


class OfficePatch(StrictModel):
    id: str
    path: str
    sha256: str = Field(pattern=r"^[0-9a-f]{64}$")
    upstream_commit: str = Field(pattern=r"^[0-9a-f]{40}$")
    upstream_change_id: str
    upstream_review: str
    license: str


class PatchSeries(StrictModel):
    series_file: str
    variant: str
    patches: list[OfficePatch]


class BuildResult(StrictModel):
    runtime_variant: str
    runtime_image: str
    identity_evidence: str
    status: str


class ConformanceDefinition(StrictModel):
    case_id: str
    fixture: str
    evidence: str
    stock_expected: Literal["failed"]
    patched_expected: Literal["passed"]


class OfficeSourceManifest(StrictModel):
    schema_version: Literal["1.0.0"]
    office_id: Literal["libreoffice"]
    upstream: UpstreamSource
    baseline: StockBaseline
    build: BuildDefinition
    patch_series: PatchSeries
    result: BuildResult
    conformance: ConformanceDefinition

    @model_validator(mode="after")
    def distinguish_patched_runtime(self) -> OfficeSourceManifest:
        if self.result.runtime_variant in {"", "stock"}:
            raise ValueError("patched runtime variant must never be stock")
        if self.result.runtime_image == self.baseline.runtime_image:
            raise ValueError("patched runtime image must differ from stock")
        if self.patch_series.variant != self.result.runtime_variant:
            raise ValueError("patch-series and result variants must match")
        return self

    def verify_files(self, repository_root: Path) -> None:
        series_path = repository_root / self.patch_series.series_file
        names = [line.strip() for line in series_path.read_text().splitlines() if line.strip()]
        expected = [Path(patch.path).name for patch in self.patch_series.patches]
        if names != expected:
            raise ValueError("patch series order does not match the manifest")
        for patch in self.patch_series.patches:
            actual = hashlib.sha256((repository_root / patch.path).read_bytes()).hexdigest()
            if actual != patch.sha256:
                raise ValueError(f"patch checksum mismatch: {patch.path}")


def load_office_source_manifest(repository_root: Path) -> OfficeSourceManifest:
    path = repository_root / "office" / "libreoffice" / "manifest.json"
    manifest = OfficeSourceManifest.model_validate(json.loads(path.read_text()))
    manifest.verify_files(repository_root)
    return manifest
