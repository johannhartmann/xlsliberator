"""Regression coverage for repaired code propagation."""

from pathlib import Path
from typing import Any

from xlsliberator.legacy_agent.agent_rewriter import (
    AgentRewriter,
    ArchitectureDesign,
    GeneratedCode,
)
from xlsliberator.validation_models import GateExecutionStatus


def test_test_and_refine_returns_and_embeds_repaired_code(tmp_path: Path, monkeypatch: Any) -> None:
    initial = GeneratedCode(
        modules={"Module.py": "def broken(:\n"},
        architecture_doc="architecture",
        completeness_score=0.5,
        known_limitations=[],
    )
    repaired = GeneratedCode(
        modules={"Module.py": "def repaired():\n    pass\n\ng_exportedScripts = (repaired,)\n"},
        architecture_doc="architecture",
        completeness_score=0.5,
        known_limitations=[],
    )
    embedded: list[dict[str, str]] = []

    class ModuleResult:
        errors = ["invalid syntax"]
        warnings: list[str] = []

    class InvalidSummary:
        syntax_errors = 1
        missing_exported_scripts = 1
        valid_syntax = 0
        total_modules = 1
        validation_details = {"Module.py": ModuleResult()}

    class ValidSummary:
        syntax_errors = 0
        missing_exported_scripts = 0
        valid_syntax = 1
        total_modules = 1
        validation_details: dict[str, Any] = {}

    summaries = iter([InvalidSummary(), ValidSummary()])
    monkeypatch.setattr(
        "xlsliberator.embed_macros.embed_python_macros",
        lambda _path, modules: embedded.append(dict(modules)),
    )
    monkeypatch.setattr(
        "xlsliberator.python_macro_manager.validate_all_embedded_macros",
        lambda _path: next(summaries),
    )
    rewriter = AgentRewriter.__new__(AgentRewriter)
    monkeypatch.setattr(rewriter, "_fix_code_errors", lambda *_args: repaired)
    architecture = ArchitectureDesign("strategy", [], [], {}, [], [])

    final_code, validation = rewriter._test_and_refine(
        initial,
        tmp_path / "book.ods",
        [],
        Any,
        architecture,
        2,
    )

    assert final_code.modules == repaired.modules
    assert embedded == [initial.modules, repaired.modules]
    assert validation.execution_successful is False
    assert validation.execution_status == GateExecutionStatus.SKIPPED
