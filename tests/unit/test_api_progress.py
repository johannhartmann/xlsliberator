import zipfile
from pathlib import Path
from typing import Any

from xlsliberator import api
from xlsliberator.ir_models import ExtractionStats, WorkbookIR


def test_convert_progress_callback_order_and_failure_tolerance(
    tmp_path: Path,
    monkeypatch: Any,
) -> None:
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    input_path.write_text("placeholder")
    events: list[str] = []

    def write_ods(_input: Path, _output: Path, **_kwargs: Any) -> None:
        with zipfile.ZipFile(output_path, "w") as archive:
            archive.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")

    monkeypatch.setattr(api, "convert_native", write_ods)
    monkeypatch.setattr(
        api,
        "extract_workbook",
        lambda path: (WorkbookIR(file_path=str(path), file_format="xlsx"), ExtractionStats()),
    )

    def callback(phase: str, _message: str, _details: dict[str, Any]) -> None:
        events.append(phase)
        if phase == "repairing":
            raise RuntimeError("callback should not crash conversion")

    report = api.convert(input_path, output_path, embed_macros=False, progress_callback=callback)

    assert report.success
    assert events[:3] == ["converting", "repairing", "analyzing"]
    assert events[-1] == "completed"
