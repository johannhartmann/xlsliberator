"""Tests for source-map-aware event binding rewrites."""

from xlsliberator.event_binding_writer import rewrite_event_bindings
from xlsliberator.validation_models import EventBindingIR, SourceRef


def test_heuristic_event_rewrite_still_works() -> None:
    """Existing Basic URL heuristic should remain compatible."""
    content = (
        'xlink:href="vnd.sun.star.script:VBAProject.Game.Start_Click'
        '?language=Basic&amp;location=document"'
    )

    updated, unresolved = rewrite_event_bindings(
        content, {"Game.bas.py": "def Start_Click(): pass"}
    )

    assert unresolved == []
    assert "Game.bas.py$Start_Click?language=Python" in updated


def test_explicit_event_binding_wins() -> None:
    """Explicit source-map event bindings should drive rewrites."""
    source = "vnd.sun.star.script:VBAProject.Game.Start?language=Basic&amp;location=document"
    target = "vnd.sun.star.script:Game.py$Start?language=Python&amp;location=document"
    binding = EventBindingIR(
        id="binding-1",
        source_ref=SourceRef(
            source_file="book.xlsm",
            artifact_type="event_binding",
            artifact_id="binding-1",
        ),
        event_name="approveAction",
        source_handler=source,
        target_script_uri=target,
    )

    updated, unresolved = rewrite_event_bindings(f'href="{source}"', {}, [binding])

    assert unresolved == []
    assert target in updated


def test_unresolved_event_binding_reported() -> None:
    """Bindings without target script URI should be returned as unresolved."""
    binding = EventBindingIR(
        id="binding-1",
        source_ref=SourceRef(
            source_file="book.xlsm",
            artifact_type="event_binding",
            artifact_id="binding-1",
        ),
        event_name="approveAction",
        source_handler="VBAProject.Game.Start",
    )

    updated, unresolved = rewrite_event_bindings("content", {}, [binding])

    assert updated == "content"
    assert unresolved == [binding]
