"""Corpus and LibreOffice build-farm MCP service contracts."""

from __future__ import annotations

import hashlib
import json
import os
import re
from pathlib import Path
from typing import Any, cast

from fastmcp import FastMCP

from xlsliberator.boundary_models import BoundaryError, BoundaryResponse, EvidenceRecord
from xlsliberator.conformance_corpus import (
    CorpusExecution,
    CorpusManifest,
    corpus_statistics,
)
from xlsliberator.container_boundary import require_application_container
from xlsliberator.demo_corpus import DemoCorpusManifest, search_demo_corpus
from xlsliberator.repair_promotion import RepairRecord, load_repair_records
from xlsliberator.validation_models import GateExecutionStatus

_ROOT = Path(__file__).resolve().parents[2]
_REPAIR_ID = re.compile(r"^[a-z0-9][a-z0-9-]+$")


def _response(
    status: GateExecutionStatus,
    *,
    data: dict[str, Any] | None = None,
    evidence: list[EvidenceRecord] | None = None,
    error: BoundaryError | None = None,
    implemented: bool = True,
    capability_available: bool = True,
) -> dict[str, Any]:
    return BoundaryResponse(
        transport_success=True,
        operation_status=status,
        implemented=implemented,
        capability_available=capability_available,
        evidence=evidence or [],
        error=error,
        data=data or {},
    ).to_payload()


def _repository_path(relative: str) -> Path:
    candidate = (_ROOT / relative).resolve()
    if not candidate.is_relative_to(_ROOT.resolve()):
        raise ValueError("path must stay inside the repository")
    return candidate


async def search_public_fixtures(query: str, limit: int = 20) -> dict[str, Any]:
    """Search redistributable public corpus and serious-episode metadata."""
    try:
        terms = {term.casefold() for term in query.split() if term.strip()}
        manifest = CorpusManifest.load(_ROOT / "corpus/manifest.json")
        fixtures = []
        for fixture in manifest.fixtures:
            haystack = " ".join(
                (
                    fixture.fixture_id,
                    fixture.format,
                    fixture.origin,
                    *fixture.categories,
                    *fixture.features,
                )
            ).casefold()
            if terms and not all(term in haystack for term in terms):
                continue
            fixtures.append(
                {
                    "fixture_id": fixture.fixture_id,
                    "format": fixture.format,
                    "origin": fixture.origin,
                    "categories": fixture.categories,
                    "features": fixture.features,
                    "path": fixture.path,
                }
            )
        demos = search_demo_corpus(
            DemoCorpusManifest.load(_ROOT / "tests/corpus/manifests/episodes.json"),
            query=query,
        )
        return _response(
            GateExecutionStatus.PASSED,
            data={
                "query": query,
                "fixtures": fixtures[: max(1, min(limit, 100))],
                "episodes": demos[: max(1, min(limit, 100))],
                "hidden_expectations_included": False,
            },
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def search_prior_failures(query: str, limit: int = 20) -> dict[str, Any]:
    """Search public repair metadata without returning private workbook content."""
    try:
        terms = {term.casefold() for term in query.split() if term.strip()}
        records = load_repair_records(_ROOT / "repairs")
        matches = [
            {
                "repair_id": record.repair_id,
                "title": record.title,
                "classification": record.classification,
                "failure_signature": record.failure_signature,
                "minimized_fixture": record.minimized_fixture.path,
                "upstream_review": record.upstream_review,
            }
            for record in records
            if not terms
            or all(
                term
                in " ".join(
                    (
                        record.repair_id,
                        record.title,
                        record.classification,
                        record.failure_signature,
                    )
                ).casefold()
                for term in terms
            )
        ]
        return _response(
            GateExecutionStatus.PASSED,
            data={"query": query, "matches": matches[: max(1, min(limit, 100))]},
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def run_public_suite(repair_id: str) -> dict[str, Any]:
    """Validate a repair's checked execution chain; never synthesize a fresh pass."""
    try:
        if _REPAIR_ID.fullmatch(repair_id) is None:
            raise ValueError("invalid repair identifier")
        record_path = _ROOT / "repairs" / repair_id / "record.json"
        if not record_path.is_file():
            return _response(
                GateExecutionStatus.UNAVAILABLE,
                capability_available=False,
                error=BoundaryError(
                    type="PublicSuiteUnavailable",
                    message="No public execution record exists for this repair",
                ),
            )
        record = RepairRecord.load(record_path)
        errors = record.verify(_ROOT)
        if errors:
            return _response(
                GateExecutionStatus.FAILED,
                data={"repair_id": repair_id, "validation_errors": errors},
                error=BoundaryError(
                    type="RepairEvidenceInvalid",
                    message="; ".join(errors),
                ),
            )
        return _response(
            GateExecutionStatus.PASSED,
            data={
                "repair_id": repair_id,
                "stock_disposition": "failed-as-expected",
                "patched_disposition": "passed",
                "target_version": "26.2.4.2",
            },
            evidence=[
                EvidenceRecord(
                    kind="repair_record",
                    path=record_path.relative_to(_ROOT).as_posix(),
                ),
                EvidenceRecord(
                    kind="stock_patched_execution",
                    path=record.exact_scenario_evidence.path,
                ),
            ],
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def register_minimized_failure(
    repair_id: str,
    failure_signature: str,
    fixture: dict[str, Any],
    provenance: str,
    license_name: str,
) -> dict[str, Any]:
    """Register a minimized public failure in a configured writable registry."""
    registry = os.environ.get("XLSLIBERATOR_CORPUS_REGISTRY_ROOT", "")
    if os.environ.get("XLSLIBERATOR_CORPUS_REGISTRATION_ENABLED", "").casefold() != "true":
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(
                type="CorpusRegistrationUnauthorized",
                message="Corpus registration is disabled for this server role",
            ),
        )
    if not registry:
        return _response(
            GateExecutionStatus.UNAVAILABLE,
            capability_available=False,
            error=BoundaryError(
                type="CorpusRegistryUnavailable",
                message="Writable corpus registry is not configured",
            ),
        )
    try:
        if _REPAIR_ID.fullmatch(repair_id) is None:
            raise ValueError("invalid repair identifier")
        payload = {
            "schema_version": "1.0.0",
            "repair_id": repair_id,
            "failure_signature": failure_signature,
            "fixture": fixture,
            "provenance": provenance,
            "license": license_name,
            "status": "registered-not-executed",
        }
        root = Path(registry).resolve()
        root.mkdir(parents=True, exist_ok=True)
        output = (root / f"{repair_id}.json").resolve()
        if not output.is_relative_to(root):
            raise ValueError("registry path escaped its configured root")
        if output.exists():
            raise FileExistsError("repair identifier is already registered")
        output.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        return _response(
            GateExecutionStatus.PASSED,
            data={
                "repair_id": repair_id,
                "status": "registered-not-executed",
                "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
            },
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def compare_runs(first_path: str, second_path: str) -> dict[str, Any]:
    """Compare two public machine-readable run records by digest and status."""
    try:
        first = _repository_path(first_path)
        second = _repository_path(second_path)
        first_payload = json.loads(first.read_text(encoding="utf-8"))
        second_payload = json.loads(second.read_text(encoding="utf-8"))
        return _response(
            GateExecutionStatus.PASSED,
            data={
                "identical": first_payload == second_payload,
                "first_sha256": hashlib.sha256(first.read_bytes()).hexdigest(),
                "second_sha256": hashlib.sha256(second.read_bytes()).hexdigest(),
                "first_status": (
                    first_payload.get("status") if isinstance(first_payload, dict) else None
                ),
                "second_status": (
                    second_payload.get("status") if isinstance(second_payload, dict) else None
                ),
            },
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def capability_report(executions_path: str) -> dict[str, Any]:
    """Generate truthful public-corpus counts from explicit execution results."""
    try:
        payload = json.loads(_repository_path(executions_path).read_text(encoding="utf-8"))
        if not isinstance(payload, dict) or not isinstance(payload.get("executions"), list):
            raise ValueError("execution result document must contain an executions list")
        executions = [
            CorpusExecution.model_validate(item)
            for item in cast(list[object], payload["executions"])
        ]
        manifest = CorpusManifest.load(_ROOT / "corpus/manifest.json")
        report = corpus_statistics(manifest, executions)
        return _response(
            GateExecutionStatus.PASSED,
            data={"report": report.model_dump(mode="json")},
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


async def run_hidden_acceptance(
    repair_id: str,
) -> dict[str, Any]:
    """Return only a sanitized hidden-suite disposition to authorized reviewers."""
    if os.environ.get("XLSLIBERATOR_HIDDEN_CORPUS_REVIEWER_ENABLED", "").casefold() != "true":
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(
                type="HiddenCorpusUnauthorized",
                message="Hidden acceptance is restricted to the reviewer service",
            ),
        )
    results_root = os.environ.get("XLSLIBERATOR_HIDDEN_CORPUS_RESULTS_ROOT", "")
    if not results_root:
        return _response(
            GateExecutionStatus.UNAVAILABLE,
            capability_available=False,
            error=BoundaryError(
                type="HiddenCorpusUnavailable",
                message="Hidden acceptance infrastructure is not configured",
            ),
        )
    try:
        if _REPAIR_ID.fullmatch(repair_id) is None:
            raise ValueError("invalid repair identifier")
        root = Path(results_root).resolve()
        result = (root / f"{repair_id}.json").resolve()
        if not result.is_relative_to(root) or not result.is_file():
            raise FileNotFoundError("hidden result is not available")
        payload = json.loads(result.read_text(encoding="utf-8"))
        if not isinstance(payload, dict) or payload.get("status") not in {
            "passed",
            "failed",
            "unavailable",
        }:
            raise ValueError("hidden result is malformed")
        return _response(
            GateExecutionStatus(str(payload["status"])),
            data={
                "repair_id": repair_id,
                "status": payload["status"],
                "sanitized_findings": payload.get("sanitized_findings", []),
                "hidden_definitions_included": False,
            },
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


def _buildfarm_unavailable(operation: str) -> dict[str, Any]:
    if os.environ.get("XLSLIBERATOR_BUILD_FARM_MUTATION_ENABLED", "").casefold() != "true":
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(
                type="BuildFarmUnauthorized",
                message=f"{operation} requires an authorized LibreOffice engineer task",
            ),
        )
    return _response(
        GateExecutionStatus.UNAVAILABLE,
        capability_available=False,
        error=BoundaryError(
            type="BuildFarmBackendUnavailable",
            message=(
                f"{operation} requires the external isolated source-build backend; "
                "the XLSLiberator application container never builds LibreOffice locally"
            ),
        ),
    )


async def create_source_worktree(source_commit: str, repair_id: str) -> dict[str, Any]:
    """Create a pinned LibreOffice source worktree in an authorized build backend."""
    del source_commit, repair_id
    return _buildfarm_unavailable("create_source_worktree")


async def apply_patch(worktree_id: str, patch_sha256: str) -> dict[str, Any]:
    """Apply one identified patch in an authorized build backend."""
    del worktree_id, patch_sha256
    return _buildfarm_unavailable("apply_patch")


async def build_component(worktree_id: str, component: str) -> dict[str, Any]:
    """Build one LibreOffice component in the external source-build backend."""
    del worktree_id, component
    return _buildfarm_unavailable("build_component")


async def run_upstream_tests(worktree_id: str, tests: list[str]) -> dict[str, Any]:
    """Run declared upstream tests in the external source-build backend."""
    del worktree_id, tests
    return _buildfarm_unavailable("run_upstream_tests")


async def publish_test_artifact(worktree_id: str, artifact_name: str) -> dict[str, Any]:
    """Publish an immutable build artifact from the external backend."""
    del worktree_id, artifact_name
    return _buildfarm_unavailable("publish_test_artifact")


async def compare_stock_patched(repair_id: str) -> dict[str, Any]:
    """Validate and return immutable stock-versus-patched repair evidence."""
    return await run_public_suite(repair_id)


async def collect_build_logs(repair_id: str) -> dict[str, Any]:
    """Return durable public build evidence paths without leaking private logs."""
    try:
        record = RepairRecord.load(_ROOT / "repairs" / repair_id / "record.json")
        errors = record.verify(_ROOT)
        if errors:
            raise ValueError("; ".join(errors))
        return _response(
            GateExecutionStatus.PASSED,
            data={
                "repair_id": repair_id,
                "evidence_paths": [
                    record.exact_scenario_evidence.path,
                    record.affected_corpus_evidence.path,
                    record.reviewer_evidence.path,
                ],
                "runtime_identity": (
                    record.libreoffice.model_dump(mode="json")
                    if record.libreoffice is not None
                    else None
                ),
            },
        )
    except Exception as exc:
        return _response(
            GateExecutionStatus.FAILED,
            error=BoundaryError(type=type(exc).__name__, message=str(exc)),
        )


corpus_mcp = FastMCP(name="XLSLiberator Migration Corpus")
for _tool in (
    search_public_fixtures,
    search_prior_failures,
    run_public_suite,
    register_minimized_failure,
    compare_runs,
    capability_report,
    run_hidden_acceptance,
):
    corpus_mcp.tool(_tool)

buildfarm_mcp = FastMCP(name="XLSLiberator LibreOffice Build Farm")
for _tool in (
    create_source_worktree,
    apply_patch,
    build_component,
    run_upstream_tests,
    publish_test_artifact,
    compare_stock_patched,
    collect_build_logs,
):
    buildfarm_mcp.tool(_tool)


def serve_corpus(host: str = "127.0.0.1", port: int = 8010) -> None:
    """Serve the corpus MCP inside the application container."""
    _serve(corpus_mcp, host, port)


def serve_buildfarm(host: str = "127.0.0.1", port: int = 8020) -> None:
    """Serve the build-farm contract MCP inside the application container."""
    _serve(buildfarm_mcp, host, port)


def _serve(server: FastMCP, host: str, port: int) -> None:
    require_application_container()
    if host not in {"127.0.0.1", "localhost", "::1"}:
        raise ValueError("trusted-local MCP may bind only to a loopback address")
    server.run(transport="http", host=host, port=port)
