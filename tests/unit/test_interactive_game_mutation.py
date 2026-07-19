"""Static safety checks for the interactive-game mutation campaign."""

import ast
import importlib.util
import sys
from pathlib import Path
from types import ModuleType

import pytest


def _load_mutation_module() -> ModuleType:
    path = Path("demos/interactive-game/mutations/run.py").resolve()
    spec = importlib.util.spec_from_file_location("interactive_game_mutation_campaign", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("mutation campaign could not be loaded")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_MUTATION_MODULE = _load_mutation_module()
MUTATIONS = _MUTATION_MODULE.MUTATIONS
MutationSpec = _MUTATION_MODULE.MutationSpec
_apply_mutation = _MUTATION_MODULE._apply_mutation


def test_every_declared_mutation_is_unique_syntactically_valid_and_applicable() -> None:
    source = Path(
        "demos/interactive-game/candidate/candidate_interactive_game/engine.py"
    ).read_text(encoding="utf-8")

    assert len(MUTATIONS) >= 5
    assert len({mutation.mutation_id for mutation in MUTATIONS}) == len(MUTATIONS)
    assert len({mutation.behavior for mutation in MUTATIONS}) == len(MUTATIONS)
    for mutation in MUTATIONS:
        mutated = _apply_mutation(source, mutation)
        assert mutated != source
        ast.parse(mutated)


def test_mutation_application_rejects_missing_or_ambiguous_match() -> None:
    missing = MutationSpec("missing", "behavior", "absent", "replacement")
    ambiguous = MutationSpec("ambiguous", "behavior", "same", "replacement")

    with pytest.raises(ValueError, match="found 0"):
        _apply_mutation("source", missing)
    with pytest.raises(ValueError, match="found 2"):
        _apply_mutation("same same", ambiguous)
