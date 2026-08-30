"""Thread-scoped deterministic tools exposed to the Open-SWE workbook agent."""

from __future__ import annotations

import asyncio
import json
import re
import shutil
import tempfile
from pathlib import Path, PurePosixPath
from typing import Any

from langchain_core.tools import BaseTool, tool
from langchain_mcp_adapters.client import MultiServerMCPClient

from xlsliberator.candidate_runtime import (
    create_candidate_workspace as create_candidate_workspace_api,
)
from xlsliberator.candidate_runtime import (
    inspect_candidate_package as inspect_candidate_package_api,
)
from xlsliberator.candidate_runtime import (
    package_candidate_workspace as package_candidate_workspace_api,
)
from xlsliberator.candidate_runtime import (
    validate_candidate_workspace as validate_candidate_workspace_api,
)
from xlsliberator.migration_check import load_acceptance
from xlsliberator.open_swe_agent.state import append_event, thread_root
from xlsliberator.xlsprobe import write_dossier

_SAFE_SCENARIO_ID = re.compile(r"^[A-Za-z0-9_.-]{1,100}$")


def workbook_tools(thread_id: str) -> list[BaseTool]:
    """Build the curated, path-confined toolset for one migration thread."""
    root = thread_root(thread_id).resolve()

    def physical_path(virtual_path: str) -> Path:
        path = PurePosixPath(virtual_path)
        try:
            relative = path.relative_to("/workspace")
        except ValueError as exc:
            raise ValueError("paths must be below /workspace") from exc
        if not relative.parts:
            raise ValueError("a workspace file path is required")
        destination = (root / Path(*relative.parts)).resolve()
        try:
            destination.relative_to(root)
        except ValueError as exc:
            raise ValueError("path escapes the migration workspace") from exc
        return destination

    def writable_path(virtual_path: str) -> Path:
        destination = physical_path(virtual_path)
        relative = destination.relative_to(root)
        protected = {"source", "dependencies", "dossier", "evidence", "migration.json"}
        if relative.parts[0] in protected:
            raise ValueError("server-owned workspace paths are read-only")
        return destination

    async def invoke(name: str, arguments: dict[str, Any]) -> dict[str, Any]:
        client = MultiServerMCPClient(
            {
                "xlsliberator": {
                    "transport": "streamable_http",
                    "url": "http://xlsliberator-mcp:8000/mcp",
                }
            }
        )
        registered = {item.name: item for item in await client.get_tools()}
        selected = registered.get(name)
        if selected is None:
            raise RuntimeError(f"required XLSLiberator MCP tool is unavailable: {name}")
        result = await selected.ainvoke(arguments)
        if isinstance(result, dict):
            return result
        text_payloads: list[str] = []
        if isinstance(result, str):
            text_payloads.append(result)
        elif isinstance(result, list):
            text_payloads.extend(
                str(block["text"])
                for block in result
                if isinstance(block, dict)
                and block.get("type") == "text"
                and isinstance(block.get("text"), str)
            )
        for text_payload in text_payloads:
            try:
                decoded = json.loads(text_payload)
            except json.JSONDecodeError:
                continue
            if isinstance(decoded, dict):
                return decoded
        raise RuntimeError(f"MCP tool {name} returned an invalid response")

    def record(name: str, payload: dict[str, Any]) -> None:
        evidence_dir = root / "evidence"
        evidence_dir.mkdir(parents=True, exist_ok=True)
        destination = evidence_dir / name
        destination.write_text(
            json.dumps(payload, indent=2, sort_keys=True, default=str),
            encoding="utf-8",
        )

    @tool
    async def create_source_dossier(
        source_path: str,
        dossier_path: str = "/workspace/dossier",
    ) -> dict[str, Any]:
        """Create a bounded dossier containing complete extracted VBA and source evidence."""
        append_event(
            thread_id,
            stage="forensics",
            message="Creating the complete bounded workbook source dossier",
        )
        source = physical_path(source_path)
        if dossier_path != "/workspace/dossier":
            raise ValueError("the source dossier path is fixed at /workspace/dossier")
        destination = physical_path(dossier_path)
        if destination.exists():
            raise ValueError("dossier destination already exists")
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = Path(tempfile.mkdtemp(prefix=".xlsprobe-agent-", dir=destination.parent))
        try:
            report = await asyncio.to_thread(write_dossier, source, temporary)
            (temporary / "migration").replace(destination)
        finally:
            shutil.rmtree(temporary, ignore_errors=True)
        files = [
            f"/workspace/{path.relative_to(root).as_posix()}"
            for path in sorted(destination.rglob("*"))
            if path.is_file()
        ]
        result = {
            "operation_status": "passed",
            "dossier_path": dossier_path,
            "source_sha256": report.source_sha256,
            "source_format": report.source_format,
            "coverage": {
                name: coverage.model_dump(mode="json")
                for name, coverage in sorted(report.coverage.items())
            },
            "evidence_index": files[:500],
            "evidence_index_truncated": len(files) > 500,
        }
        record("source-dossier.json", result)
        return result

    @tool
    async def create_candidate_workspace(
        source_path: str,
        workspace_path: str,
        candidate_id: str,
    ) -> dict[str, Any]:
        """Scaffold a source-bound workspace where target-native candidate code is authored."""
        result = await asyncio.to_thread(
            create_candidate_workspace_api,
            physical_path(source_path),
            writable_path(workspace_path),
            candidate_id=candidate_id,
        )
        return result

    @tool
    async def validate_candidate_workspace(
        source_path: str,
        workspace_path: str,
    ) -> dict[str, Any]:
        """Validate candidate paths, entrypoints and a freshly calculated source binding."""
        return await asyncio.to_thread(
            validate_candidate_workspace_api,
            physical_path(source_path),
            physical_path(workspace_path),
        )

    @tool
    async def package_candidate_workspace(
        source_path: str,
        workspace_path: str,
        candidate_path: str,
    ) -> dict[str, Any]:
        """Create a deterministic content-bound candidate ZIP and return its hash."""
        candidate = writable_path(candidate_path)
        candidate.parent.mkdir(parents=True, exist_ok=True)
        result = await asyncio.to_thread(
            package_candidate_workspace_api,
            physical_path(source_path),
            physical_path(workspace_path),
            candidate,
        )
        record("candidate-package.json", result)
        return result

    @tool
    async def inspect_candidate_package(candidate_path: str) -> dict[str, Any]:
        """Inspect a packaged candidate's hashes and manifest without executing its code."""
        return await asyncio.to_thread(
            inspect_candidate_package_api,
            physical_path(candidate_path),
        )

    @tool
    async def validate_acceptance_scenario(acceptance_path: str) -> dict[str, Any]:
        """Validate an authored public acceptance YAML/JSON before runtime execution."""
        acceptance = await asyncio.to_thread(load_acceptance, physical_path(acceptance_path))
        required_observations = [
            request
            for step in acceptance.scenario.steps
            for request in (*step.observations_before, *step.observations_after)
            if request.required
        ]
        expected_observations = [
            request for request in required_observations if request.expected is not None
        ]
        if not expected_observations:
            raise ValueError(
                "acceptance must contain at least one required authored expected observation"
            )
        return {
            "operation_status": "passed",
            "migration_id": acceptance.migration.id,
            "scenario_id": acceptance.scenario.id,
            "step_count": len(acceptance.scenario.steps),
            "required_observation_count": len(required_observations),
            "expected_observation_count": len(expected_observations),
        }

    @tool
    async def run_libreoffice_scenario(
        target_path: str,
        acceptance_path: str,
    ) -> dict[str, Any]:
        """Run authored expectations in pinned LibreOffice and persist typed evidence."""
        append_event(
            thread_id,
            stage="validation",
            message="Running the public LibreOffice acceptance scenario",
        )
        evidence = root / "evidence" / "public-acceptance"
        evidence.parent.mkdir(parents=True, exist_ok=True)
        result = await invoke(
            "run_declared_acceptance",
            {
                "target_path": str(physical_path(target_path)),
                "acceptance_path": str(physical_path(acceptance_path)),
                "evidence_path": str(evidence),
            },
        )
        record("public-acceptance.json", result)
        return result

    @tool
    async def run_mutation_suite(
        target_path: str,
        acceptance_path: str,
    ) -> dict[str, Any]:
        """Mutate target formulas/code and require the public acceptance to kill mutants."""
        append_event(
            thread_id,
            stage="validation",
            message="Running target-native mutation testing",
        )
        evidence = root / "evidence" / "mutations"
        evidence.parent.mkdir(parents=True, exist_ok=True)
        result = await invoke(
            "run_declared_mutation_suite",
            {
                "target_path": str(physical_path(target_path)),
                "acceptance_path": str(physical_path(acceptance_path)),
                "evidence_path": str(evidence),
            },
        )
        record("mutation-suite.json", result)
        return result

    @tool
    async def run_gui_scenario(
        target_path: str,
        candidate_path: str,
        scenario_id: str,
        actions: list[dict[str, Any]],
        adapter_config: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Run declared real-X11 control events and screenshots in the pinned GUI runtime."""
        if _SAFE_SCENARIO_ID.fullmatch(scenario_id) is None:
            raise ValueError("scenario_id is malformed")
        append_event(
            thread_id,
            stage="validation",
            message="Running the real LibreOffice GUI acceptance scenario",
        )
        evidence = root / "evidence" / "gui" / f"{scenario_id}.zip"
        evidence.parent.mkdir(parents=True, exist_ok=True)
        result = await invoke(
            "run_application_scenario",
            {
                "target_path": str(physical_path(target_path)),
                "candidate_path": str(physical_path(candidate_path)),
                "evidence_path": str(evidence),
                "scenario_id": scenario_id,
                "actions": actions,
                "adapter_config": adapter_config or {},
            },
        )
        record(f"gui-{scenario_id}.json", result)
        return result

    @tool
    async def convert_baseline_to_ods(source_path: str, target_path: str) -> dict[str, Any]:
        """Create a deterministic baseline ODS before agent-authored repairs."""
        append_event(
            thread_id,
            stage="libreoffice",
            message="Creating the deterministic LibreOffice baseline",
        )
        target = writable_path(target_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        result = await invoke(
            "convert_excel_to_ods",
            {
                "excel_path": str(physical_path(source_path)),
                "output_path": str(target),
                "embed_macros": False,
            },
        )
        record("conversion.json", result)
        return result

    @tool
    async def build_candidate(
        source_path: str,
        candidate_path: str,
        target_path: str,
    ) -> dict[str, Any]:
        """Build an ODS from an agent-authored, content-bound Python/UNO candidate."""
        append_event(
            thread_id,
            stage="specialists",
            message="Building the target-native migration candidate",
        )
        target = writable_path(target_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        result = await invoke(
            "build_application_candidate",
            {
                "source_path": str(physical_path(source_path)),
                "candidate_path": str(physical_path(candidate_path)),
                "output_path": str(target),
            },
        )
        record("candidate-build.json", result)
        return result

    @tool
    async def certify_transformation(source_path: str, target_path: str) -> dict[str, Any]:
        """Run the fail-closed XLSLiberator certification gates."""
        append_event(
            thread_id,
            stage="validation",
            message="Running deterministic transformation certification",
        )
        result = await invoke(
            "validate_transformation",
            {
                "excel_path": str(physical_path(source_path)),
                "ods_path": str(physical_path(target_path)),
                "target": "libreoffice",
            },
        )
        record("certification.json", result)
        return result

    @tool
    async def verify_save_close_reopen(target_path: str) -> dict[str, Any]:
        """Open the ODS in pinned LibreOffice, recalculate, save, close and reopen it."""
        append_event(
            thread_id,
            stage="libreoffice",
            message="Verifying save, close and reopen in LibreOffice 26.2.4.2",
        )
        target = str(writable_path(target_path))
        steps: dict[str, Any] = {}
        created = await invoke("create_session", {"environment": {}})
        steps["create_session"] = created
        session_id = created.get("session_id")
        if not isinstance(session_id, str) or not session_id:
            record("save-reopen.json", {"operation_status": "failed", "steps": steps})
            return {"operation_status": "failed", "steps": steps}
        try:
            for name, arguments in (
                ("open_document", {"session_id": session_id, "document_path": target}),
                ("inspect_document", {"session_id": session_id}),
                ("recalculate", {"session_id": session_id}),
                ("save", {"session_id": session_id, "output_path": target}),
                ("close", {"session_id": session_id}),
                ("reopen", {"session_id": session_id}),
                ("inspect_document", {"session_id": session_id}),
                ("collect_logs", {"session_id": session_id}),
            ):
                steps[f"{name}-{len(steps)}"] = await invoke(name, arguments)
        finally:
            steps["destroy_session"] = await invoke(
                "destroy_session",
                {"session_id": session_id},
            )
        passed = all(
            isinstance(result, dict) and result.get("operation_status") == "passed"
            for result in steps.values()
        )
        payload = {
            "operation_status": "passed" if passed else "failed",
            "target_libreoffice_version": "26.2.4.2",
            "steps": steps,
        }
        record("save-reopen.json", payload)
        return payload

    return [
        create_source_dossier,
        create_candidate_workspace,
        validate_candidate_workspace,
        package_candidate_workspace,
        inspect_candidate_package,
        validate_acceptance_scenario,
        run_libreoffice_scenario,
        run_mutation_suite,
        run_gui_scenario,
        convert_baseline_to_ods,
        build_candidate,
        certify_transformation,
        verify_save_close_reopen,
    ]
