"""Upload validation and filesystem safety helpers."""

from __future__ import annotations

import re
import uuid
from dataclasses import dataclass
from pathlib import Path

ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".xlsm", ".xlsb"}
ZIP_SIGNATURE = b"PK\x03\x04"
OLE_CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


class UploadValidationError(ValueError):
    """Raised when an uploaded file fails validation."""


@dataclass(frozen=True)
class JobPaths:
    """Internal paths for a conversion job."""

    job_dir: Path
    input_path: Path
    output_path: Path
    report_json_path: Path
    report_md_path: Path
    log_bundle_path: Path
    profile_dir: Path


def validate_upload_filename(filename: str) -> str:
    """Return a normalized extension or raise for unsupported/dangerous names."""
    raw = filename.strip()
    if not raw:
        raise UploadValidationError("Missing filename")
    name = Path(raw).name
    if name != raw or "/" in raw or "\\" in raw or "\x00" in raw:
        raise UploadValidationError("Filename must not contain path components")
    suffix = Path(name).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise UploadValidationError("Unsupported upload extension")
    if _has_dangerous_double_extension(name, suffix):
        raise UploadValidationError("Dangerous double extension")
    return suffix


def generate_job_id() -> str:
    """Return a server-generated job identifier."""
    return str(uuid.uuid4())


def safe_job_paths(data_dir: Path, job_id: str, original_filename: str) -> JobPaths:
    """Return server-controlled paths for a job without reusing uploaded names."""
    if not _UUID_RE.fullmatch(job_id):
        raise UploadValidationError("Invalid job id")
    extension = validate_upload_filename(original_filename)
    jobs_root = data_dir / "jobs"
    job_dir = (jobs_root / job_id).resolve()
    root = jobs_root.resolve()
    if root not in job_dir.parents:
        raise UploadValidationError("Invalid job path")
    return JobPaths(
        job_dir=job_dir,
        input_path=job_dir / f"input{extension}",
        output_path=job_dir / "output.ods",
        report_json_path=job_dir / "report.json",
        report_md_path=job_dir / "report.md",
        log_bundle_path=job_dir / "logs.zip",
        profile_dir=job_dir / "lo-profile",
    )


def detect_basic_signature(path: Path) -> str:
    """Detect the coarse spreadsheet container signature."""
    with path.open("rb") as handle:
        header = handle.read(8)
    if header.startswith(ZIP_SIGNATURE):
        return "xlsx_zip"
    if header.startswith(OLE_CFB_SIGNATURE):
        return "ole_cfb"
    return "unknown"


def validate_upload_signature(path: Path, extension: str) -> None:
    """Validate extension against a minimal file signature allowlist."""
    signature = detect_basic_signature(path)
    if extension in {".xlsx", ".xlsm"} and signature != "xlsx_zip":
        raise UploadValidationError("Uploaded OOXML workbook is not a ZIP container")
    if extension == ".xls" and signature != "ole_cfb":
        raise UploadValidationError("Uploaded .xls workbook is not an OLE CFB container")
    if extension == ".xlsb" and signature not in {"xlsx_zip", "ole_cfb"}:
        raise UploadValidationError("Uploaded .xlsb workbook has an unsupported container")


def safe_download_stem(original_filename: str) -> str:
    """Return a conservative user-facing download stem."""
    stem = Path(original_filename).stem
    clean = re.sub(r"[^A-Za-z0-9._-]+", "-", stem).strip(".-_")
    return clean or "workbook"


def _has_dangerous_double_extension(filename: str, suffix: str) -> bool:
    lower = filename.lower()
    dangerous = {".exe", ".bat", ".cmd", ".com", ".js", ".sh", ".php", ".html"}
    without_suffix = lower[: -len(suffix)]
    return any(without_suffix.endswith(ext) for ext in dangerous)


_UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
