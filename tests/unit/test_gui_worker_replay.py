"""Office-free checks for the portable GUI replay artifact."""

from pathlib import Path

from xlsliberator.gui_worker import _write_replay_html


def test_replay_html_embeds_timeline_without_interpreting_result_markup(tmp_path: Path) -> None:
    response = {
        "operations": [
            {
                "sequence": 1,
                "kind": "key",
                "status": "passed",
                "duration_ms": 125,
                "result": {"value": "</script><script>alert(1)</script>"},
            }
        ]
    }

    _write_replay_html(tmp_path, response)

    replay = (tmp_path / "replay.html").read_text(encoding="utf-8")
    assert 'src="recording.webm"' in replay
    assert "evidence.operations.forEach" in replay
    assert "</script><script>alert(1)</script>" not in replay
    assert "\\u003c/script>" in replay
