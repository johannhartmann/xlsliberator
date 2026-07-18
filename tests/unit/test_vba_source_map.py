"""Tests for VBA source-map markers."""


def test_inject_source_markers_emits_stable_marker_format() -> None:
    """Injected markers carry the module, procedure, and artifact id."""
    from xlsliberator.legacy_agent.vba2py_uno import _inject_source_markers

    code = "def StartButton_Click():\n    pass\n"
    updated = _inject_source_markers(code, ["StartButton_Click"], "Game.bas")

    assert "# xlsliberator-source:" in updated
    assert "module=Game.bas" in updated
    assert "procedure=StartButton_Click" in updated
    assert "artifact_id=Game.bas.StartButton_Click" in updated


def test_collect_source_map_markers() -> None:
    """Macro enumeration helpers should collect source map comments."""
    from xlsliberator.python_macro_manager import collect_source_map_markers

    markers = collect_source_map_markers(
        """
def Start():
    # xlsliberator-source: module=Game; procedure=Start; artifact_id=Game.Start
    pass
"""
    )

    assert markers == [
        "# xlsliberator-source: module=Game; procedure=Start; artifact_id=Game.Start"
    ]


def test_extract_vba_procedure_names_ignores_end_and_exit() -> None:
    """End Sub / Exit Sub and the following line must not be captured as procedures.

    These helpers feed the LLM translation path's source-map injection.
    """
    from xlsliberator.legacy_agent.vba2py_uno import _extract_vba_procedure_names

    code = "Private Sub Foo()\n    Exit Sub\nEnd Sub\nPublic Function Bar()\nEnd Function\n"

    assert _extract_vba_procedure_names(code) == ["Foo", "Bar"]


def test_inject_source_markers_handles_nested_parens() -> None:
    """A signature with nested parens in a default value still gets a source marker."""
    from xlsliberator.legacy_agent.vba2py_uno import _inject_source_markers

    updated = _inject_source_markers("def foo(x=(1, 2)):\n    return x\n", ["foo"], "Mod1")

    lines = updated.splitlines()
    assert lines[0] == "def foo(x=(1, 2)):"
    assert lines[1].strip().startswith("# xlsliberator-source: module=Mod1; procedure=foo")
