"""Open-SWE graph customized for autonomous Excel-to-LibreOffice migration."""

from __future__ import annotations

import os

from agent.middleware import SanitizeToolInputsMiddleware, ToolErrorMiddleware
from agent.utils.deferred_model import make_deferred_error_model
from agent.utils.model import make_model, provider_model_kwargs
from agent.utils.tracing import traced_graph_factory
from deepagents import create_deep_agent
from deepagents.backends import CompositeBackend, FilesystemBackend, StateBackend
from deepagents.middleware.subagents import SubAgent
from langchain_core.language_models import BaseChatModel
from langgraph.graph.state import RunnableConfig
from langgraph.pregel import Pregel

from xlsliberator.open_swe_agent.state import append_event, thread_root
from xlsliberator.open_swe_agent.tools import workbook_tools

SYSTEM_PROMPT = """You are the XLSLiberator Open-SWE migration lead.

You migrate the workbook in /workspace/source directly to LibreOffice 26.2.4.2.
LibreOffice is the only target. Never create an Excel runtime, VBA interpreter,
Excel object-model facade, compatibility VM, or deterministic migration
orchestrator. Generated behavior must be target-native Python/UNO, ODF, controls,
extensions, or open services.

Use the source-forensics tool before changing anything. Delegate focused analysis
to the forensics, VBA-liberation, and review subagents. Create a deterministic
baseline, repair it when source behavior requires generated code, then run
certification and save/close/reopen verification. Treat every failed, skipped,
unavailable, or not-run required operation as blocking.

Write deliverables below /workspace/deliverables. Completion requires:
- /workspace/deliverables/target.ods
- /workspace/deliverables/report.json and report.md
- /workspace/evidence/reviewer.json with decision APPROVE
- passed deterministic certification and save/reopen evidence

Never invent evidence, weaken validation, expose hidden reasoning, or claim that
an unavailable UI/runtime operation passed. Do not call GitHub, create commits,
push branches, or open pull requests for a workbook migration.
"""


def _specialist(name: str, description: str, system_prompt: str, model: BaseChatModel) -> SubAgent:
    return {
        "name": name,
        "description": description,
        "system_prompt": system_prompt,
        "model": model,
    }


def _configured_model() -> BaseChatModel:
    model_id = os.environ.get("XLSLIBERATOR_OPEN_SWE_MODEL", "").strip()
    if not model_id:
        return make_deferred_error_model(
            ValueError(
                "XLSLIBERATOR_OPEN_SWE_MODEL is required; no model is selected automatically"
            )
        )
    if model_id.startswith("github_models:"):
        if os.environ.get("XLSLIBERATOR_GITHUB_MODELS_ENABLED") != "1":
            return make_deferred_error_model(
                ValueError("GitHub Models requires XLSLIBERATOR_GITHUB_MODELS_ENABLED=1"),
                model_id=model_id,
            )
        token = os.environ.get("GITHUB_MODELS_TOKEN", "").strip()
        if not token:
            return make_deferred_error_model(
                ValueError("GITHUB_MODELS_TOKEN is required for an explicit GitHub Models run"),
                model_id=model_id,
            )
        from langchain_openai import ChatOpenAI

        return ChatOpenAI(
            api_key=token,
            base_url="https://models.github.ai/inference",
            model=model_id.removeprefix("github_models:"),
            max_tokens=_max_tokens(),
            max_retries=2,
        )

    from agent.dashboard.options import SUPPORTED_MODEL_IDS

    if model_id not in SUPPORTED_MODEL_IDS:
        return make_deferred_error_model(
            ValueError(f"unsupported Open-SWE model: {model_id}"),
            model_id=model_id,
        )
    kwargs = provider_model_kwargs(
        model_id,
        os.environ.get("XLSLIBERATOR_OPEN_SWE_REASONING_EFFORT", "medium"),
        max_tokens=_max_tokens(),
    )
    try:
        return make_model(model_id, use_gateway=False, **kwargs)
    except Exception as exc:
        return make_deferred_error_model(exc, model_id=model_id)


def _max_tokens() -> int:
    raw = os.environ.get("XLSLIBERATOR_OPEN_SWE_MAX_OUTPUT_TOKENS", "32768")
    try:
        return max(1024, min(int(raw), 131072))
    except ValueError:
        return 32768


async def get_agent(config: RunnableConfig) -> Pregel:
    """Build one thread-confined Open-SWE graph without a shell backend."""
    configurable = config.get("configurable")
    thread_id = configurable.get("thread_id") if isinstance(configurable, dict) else None
    if not isinstance(thread_id, str) or not thread_id:
        raise ValueError("Open-SWE migration graph requires a thread_id")

    root = thread_root(thread_id)
    root.mkdir(parents=True, exist_ok=True)
    append_event(
        thread_id,
        stage="lead",
        message="Open-SWE migration lead started",
    )
    model = _configured_model()
    tools = workbook_tools(thread_id)
    backend = CompositeBackend(
        default=StateBackend(),
        routes={
            "/workspace/": FilesystemBackend(
                root_dir=str(root),
                virtual_mode=True,
            )
        },
    )
    subagents: list[SubAgent] = [
        _specialist(
            "workbook-forensics",
            "Analyzes source-derived workbook inventories and identifies behavioral risks.",
            "Read the inventory and source-derived evidence. Report formulas, VBA, events, "
            "controls, dependencies, unknowns, and acceptance scenarios. Do not edit files.",
            model,
        ),
        _specialist(
            "vba-liberation",
            "Designs direct target-native replacements for VBA and proprietary dependencies.",
            "Design Python/UNO, ODF, control, extension, or open-service replacements. "
            "Never propose a VBA runtime, Excel facade, or Windows/Office dependency.",
            model,
        ),
        _specialist(
            "independent-review",
            "Reviews the completed migration evidence and returns APPROVE, REVISE, or BLOCK.",
            "Independently inspect source inventory, target artifacts, certification, and "
            "save/reopen evidence. Write /workspace/evidence/reviewer.json. APPROVE only "
            "when required behavior and dependency-liberation evidence genuinely pass.",
            model,
        ),
    ]
    return create_deep_agent(
        model=model,
        system_prompt=SYSTEM_PROMPT,
        tools=tools,
        subagents=subagents,
        backend=backend,
        middleware=[
            SanitizeToolInputsMiddleware(),
            ToolErrorMiddleware(),
        ],
    ).with_config(config)


traced_agent = traced_graph_factory(get_agent, "xlsliberator-open-swe")
