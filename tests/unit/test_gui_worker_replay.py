"""Office-free checks for the portable GUI replay artifact."""

from pathlib import Path

import pytest

from xlsliberator.gui_worker import _confined_path, _write_replay_html


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


def test_gui_worker_separates_read_only_inputs_from_job_outputs(tmp_path: Path) -> None:
    input_root = tmp_path / "input"
    job_root = tmp_path / "job"
    input_root.mkdir()
    job_root.mkdir()
    source = input_root / "target.ods"
    source.write_bytes(b"ods")

    assert (
        _confined_path(
            source,
            root=input_root,
            must_exist=True,
            label="input",
        )
        == source
    )
    assert (
        _confined_path(
            job_root / "evidence.zip",
            root=job_root,
            must_exist=False,
            label="output",
        )
        == job_root / "evidence.zip"
    )
    with pytest.raises(ValueError, match="must remain inside"):
        _confined_path(
            source,
            root=job_root,
            must_exist=True,
            label="output",
        )
