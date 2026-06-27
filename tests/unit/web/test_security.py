from pathlib import Path

import pytest

from xlsliberator.web.security import (
    UploadValidationError,
    detect_basic_signature,
    generate_job_id,
    safe_download_stem,
    safe_job_paths,
    validate_upload_filename,
    validate_upload_signature,
)


def test_validate_upload_filename_allows_supported_extensions() -> None:
    assert validate_upload_filename("Workbook.XLSX") == ".xlsx"
    assert validate_upload_filename("book.xlsm") == ".xlsm"
    assert validate_upload_filename("book.xlsb") == ".xlsb"
    assert validate_upload_filename("book.xls") == ".xls"


@pytest.mark.parametrize("filename", ["../book.xlsx", "book.xlsx.exe", "book.csv", ""])
def test_validate_upload_filename_rejects_unsafe_names(filename: str) -> None:
    with pytest.raises(UploadValidationError):
        validate_upload_filename(filename)


def test_safe_job_paths_uses_server_controlled_names(tmp_path: Path) -> None:
    job_id = generate_job_id()
    paths = safe_job_paths(tmp_path, job_id, "Quarterly Report.xlsx")

    assert paths.job_dir == (tmp_path / "jobs" / job_id).resolve()
    assert paths.input_path.name == "input.xlsx"
    assert paths.output_path.name == "output.ods"
    assert "Quarterly" not in str(paths.input_path)


def test_signature_detection_and_validation(tmp_path: Path) -> None:
    xlsx = tmp_path / "book.xlsx"
    xlsx.write_bytes(b"PK\x03\x04rest")
    xls = tmp_path / "book.xls"
    xls.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1rest")

    assert detect_basic_signature(xlsx) == "xlsx_zip"
    assert detect_basic_signature(xls) == "ole_cfb"
    validate_upload_signature(xlsx, ".xlsx")
    validate_upload_signature(xls, ".xls")


def test_signature_validation_rejects_mismatch(tmp_path: Path) -> None:
    path = tmp_path / "bad.xlsx"
    path.write_bytes(b"not a spreadsheet")

    with pytest.raises(UploadValidationError):
        validate_upload_signature(path, ".xlsx")


def test_safe_download_stem() -> None:
    assert safe_download_stem("Q4 report.xlsx") == "Q4-report"
