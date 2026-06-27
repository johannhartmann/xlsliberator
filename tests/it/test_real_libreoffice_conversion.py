"""Real LibreOffice conversion integration tests."""

import subprocess
from pathlib import Path

import openpyxl
import pytest

from xlsliberator.api import convert


@pytest.mark.integration
def test_convert_xlsx_to_ods_reopens_in_real_libreoffice(
    tmp_path: Path,
    skip_if_no_lo: None,
) -> None:
    """Convert a generated XLSX and verify LibreOffice recalculates the ODS."""
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.ods"
    roundtrip_dir = tmp_path / "roundtrip"
    roundtrip_dir.mkdir()

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sheet1"
    sheet["A1"] = 2
    sheet["A2"] = 3
    sheet["A3"] = "=SUM(A1:A2)"
    sheet["B1"] = "real"
    workbook.save(input_path)
    workbook.close()

    report = convert(input_path, output_path, embed_macros=False, use_agent=False)

    assert report.success
    assert output_path.exists()

    export = subprocess.run(
        [
            "soffice",
            "--headless",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(roundtrip_dir),
            str(output_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
        check=False,
    )

    assert export.returncode == 0, export.stderr or export.stdout
    roundtrip_path = roundtrip_dir / "output.xlsx"
    assert roundtrip_path.exists()

    roundtrip = openpyxl.load_workbook(roundtrip_path, data_only=True)
    try:
        assert roundtrip["Sheet1"]["A3"].value == 5
        assert roundtrip["Sheet1"]["B1"].value == "real"
    finally:
        roundtrip.close()
