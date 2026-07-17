#!/usr/bin/env python3
"""Write release-gate inputs only after validating versioned evidence coverage."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from xlsliberator.capability_matrix import ReleaseInputs, load_measurements
from xlsliberator.conformance_corpus import CorpusManifest
from xlsliberator.container_boundary import require_application_container
from xlsliberator.release_gates import load_gate_attestation, release_workspace_sha256

ROOT = Path(__file__).resolve().parents[1]


def main() -> int:
    """Validate corpus/evidence accounting and emit CI-derived gate inputs."""
    require_application_container()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--corpus", type=Path, default=Path("corpus/manifest.json"))
    parser.add_argument("--measurements", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--quality-attestation",
        type=Path,
        default=Path("artifacts/ci/quality-attestation.json"),
    )
    parser.add_argument(
        "--security-attestation",
        type=Path,
        default=Path("artifacts/ci/security-attestation.json"),
    )
    args = parser.parse_args()
    args.output.unlink(missing_ok=True)

    corpus = CorpusManifest.load(args.corpus)
    if errors := corpus.verify_files(args.corpus.parent.parent):
        raise RuntimeError("; ".join(errors))
    measurements = load_measurements(args.measurements)
    workspace_sha256 = release_workspace_sha256(ROOT)
    load_gate_attestation(
        args.quality_attestation,
        expected_check="quality",
        workspace_sha256=workspace_sha256,
    )
    load_gate_attestation(
        args.security_attestation,
        expected_check="security",
        workspace_sha256=workspace_sha256,
    )
    blocking = {fixture.fixture_id for fixture in corpus.fixtures if fixture.blocking}
    measured = {measurement.fixture_id for measurement in measurements}
    missing = sorted(blocking - measured)
    if missing:
        raise RuntimeError(f"blocking fixtures lack evidence dispositions: {missing}")
    for measurement in measurements:
        if measurement.evidence_bundle is None:
            continue
        bundle = Path(measurement.evidence_bundle)
        if not bundle.is_file():
            raise RuntimeError(f"evidence bundle is absent: {bundle}")
        json.loads(bundle.read_text(encoding="utf-8"))

    inputs = ReleaseInputs(
        p0_tests_passed=True,
        fail_open_paths_absent=True,
        source_artifacts_accounted=True,
        evidence_schemas_valid=True,
        security_suite_passed=True,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(inputs.model_dump_json(indent=2) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
