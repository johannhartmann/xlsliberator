"""Thread-scoped deterministic tools exposed to the Open-SWE workbook agent."""

from __future__ import annotations

import json
from pathlib import Path, PurePosixPath
from typing import Any

from langchain_core.tools import BaseTool, tool
from langchain_mcp_adapters.client import MultiServerMCPClient

from xlsliberator.open_swe_agent.state import append_event, thread_root


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
    async def inspect_source_workbook(source_path: str) -> dict[str, Any]:
        """Inspect an Excel workbook and persist its source-derived inventory."""
        append_event(
            thread_id,
            stage="forensics",
            message="Inspecting workbook structure, formulas, VBA, controls and dependencies",
        )
        result = await invoke(
            "inspect_workbook",
            {"excel_path": str(physical_path(source_path))},
        )
        record("source-inventory.json", result)
        return result

    @tool
    async def convert_baseline_to_ods(source_path: str, target_path: str) -> dict[str, Any]:
        """Create a deterministic baseline ODS before agent-authored repairs."""
        append_event(
            thread_id,
            stage="libreoffice",
            message="Creating the deterministic LibreOffice baseline",
        )
        target = physical_path(target_path)
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
    async def build_generated_candidate(
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
        target = physical_path(target_path)
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
        target = str(physical_path(target_path))
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
        inspect_source_workbook,
        convert_baseline_to_ods,
        build_generated_candidate,
        certify_transformation,
        verify_save_close_reopen,
    ]
