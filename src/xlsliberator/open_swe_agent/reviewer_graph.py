"""Fresh-context, read-only reviewer graph for completed migrations."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Literal

from agent.utils.tracing import traced_graph_factory
from deepagents import create_deep_agent
from deepagents.backends import CompositeBackend, FilesystemBackend, StateBackend
from deepagents.backends.protocol import DeleteResult, EditResult, WriteResult
from langchain_core.tools import tool
from langgraph.graph.state import RunnableConfig
from langgraph.pregel import Pregel

from xlsliberator.open_swe_agent.graph import _configured_model
from xlsliberator.open_swe_agent.state import workspace_root

REVIEWER_PROMPT = """You are the independent XLSLiberator migration reviewer.

You have a fresh model context and read-only access to a server-created snapshot.
Inspect review-context.json, the exact source dossier, target artifact, authored
acceptance, certification, LibreOffice execution evidence, mutation evidence, and
save/reopen evidence. Do not edit artifacts. Call submit_review exactly once with
APPROVE only when the target-native migration and every feature-appropriate required
test genuinely pass. Use REVISE for correctable gaps and BLOCK for unsafe or
fundamentally unsupported results. Never infer success from missing evidence.
"""


class ReadOnlyFilesystemBackend(FilesystemBackend):
    """Filesystem backend that permits inspection but rejects every mutation."""

    def write(self, file_path: str, content: str) -> WriteResult:
        del content
        return WriteResult(error="review workspace is read-only", path=file_path)

    def edit(
        self,
        file_path: str,
        old_string: str,
        new_string: str,
        replace_all: bool = False,
    ) -> EditResult:
        del old_string, new_string, replace_all
        return EditResult(error="review workspace is read-only", path=file_path)

    def delete(self, file_path: str) -> DeleteResult:
        return DeleteResult(error="review workspace is read-only", path=file_path)


def _review_root(review_id: str) -> Path:
    if not review_id or any(character not in "0123456789abcdef-" for character in review_id):
        raise ValueError("invalid review identifier")
    return workspace_root() / ".reviews" / review_id


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


async def get_reviewer(config: RunnableConfig) -> Pregel:
    """Build a reviewer with a distinct model and no production-edit tools."""
    configurable = config.get("configurable")
    review_id = configurable.get("review_id") if isinstance(configurable, dict) else None
    if not isinstance(review_id, str) or not review_id:
        raise ValueError("reviewer graph requires a review_id")
    root = _review_root(review_id).resolve(strict=True)
    context = json.loads((root / "review-context.json").read_text(encoding="utf-8"))
    reviewer_model = os.environ.get("XLSLIBERATOR_REVIEWER_MODEL", "").strip()
    migration_model = os.environ.get("XLSLIBERATOR_OPEN_SWE_MODEL", "").strip()
    if not reviewer_model:
        raise ValueError("XLSLIBERATOR_REVIEWER_MODEL is required")
    if reviewer_model == migration_model:
        raise ValueError("reviewer model must differ from the migration model")
    model = _configured_model(model_env="XLSLIBERATOR_REVIEWER_MODEL")

    @tool
    def submit_review(
        decision: Literal["APPROVE", "REVISE", "BLOCK"],
        rationale: str,
    ) -> dict[str, str]:
        """Submit the sole reviewer decision through a server-owned write path."""
        destination = root / "review-decision.json"
        if destination.exists():
            raise ValueError("a reviewer decision was already submitted")
        source = root / str(context["source_path"])
        target = root / str(context["target_path"])
        record = {
            "schema_version": "1.0.0",
            "review_id": review_id,
            "decision": decision,
            "rationale": rationale[:4000],
            "reviewer_model": reviewer_model,
            "source_sha256": _hash_file(source),
            "candidate_sha256": _hash_file(target),
            "test_set_identity": str(context["test_set_identity"]),
        }
        destination.write_text(
            json.dumps(record, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        return {"decision": decision, "review_id": review_id}

    backend = CompositeBackend(
        default=StateBackend(),
        routes={
            "/workspace/": ReadOnlyFilesystemBackend(
                root_dir=str(root),
                virtual_mode=True,
            )
        },
    )
    return create_deep_agent(
        model=model,
        system_prompt=REVIEWER_PROMPT,
        tools=[submit_review],
        backend=backend,
    ).with_config(config)


traced_reviewer = traced_graph_factory(get_reviewer, "xlsliberator-independent-review")
