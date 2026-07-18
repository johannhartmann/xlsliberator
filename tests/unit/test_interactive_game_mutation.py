"""Static safety checks for the interactive-game mutation campaign."""

import ast
from pathlib import Path

import pytest

from xlsliberator.interactive_game_mutation import MUTATIONS, MutationSpec, _apply_mutation


def test_every_declared_mutation_is_unique_syntactically_valid_and_applicable() -> None:
    source = Path("src/xlsliberator/interactive_game_engine.py").read_text(encoding="utf-8")

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
