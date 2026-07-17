"""Fail-closed structural limits for untrusted workbook inputs."""

from __future__ import annotations

import re
import stat
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath


class UnsafeWorkbookError(ValueError):
    """A workbook violates deterministic untrusted-input limits."""


@dataclass(frozen=True)
class WorkbookInputLimits:
    max_archive_bytes: int = 256 * 1024**2
    max_entries: int = 20_000
    max_uncompressed_bytes: int = 1024 * 1024**2
    max_part_bytes: int = 128 * 1024**2
    max_compression_ratio: float = 200.0
    max_formula_characters: int = 16_384
    max_macro_characters: int = 2_000_000


_FORMULA_PATTERN = re.compile(rb"<[^>]*(?:f|formula)[^>]*>(.*?)</[^>]+>", re.DOTALL)
_MACRO_SUFFIXES = (".bas", ".cls", ".frm", ".py")


def validate_untrusted_workbook(path: Path, limits: WorkbookInputLimits | None = None) -> None:
    """Reject hostile package structure before parser or runtime execution."""

    limits = limits or WorkbookInputLimits()
    size = path.stat().st_size
    if size > limits.max_archive_bytes:
        raise UnsafeWorkbookError("workbook exceeds the compressed input-size limit")
    if not zipfile.is_zipfile(path):
        return
    total = 0
    with zipfile.ZipFile(path) as archive:
        entries = archive.infolist()
        if len(entries) > limits.max_entries:
            raise UnsafeWorkbookError("workbook package contains too many parts")
        for info in entries:
            _validate_member_name(info.filename)
            mode = info.external_attr >> 16
            if stat.S_ISLNK(mode):
                raise UnsafeWorkbookError(f"workbook package contains a symlink: {info.filename}")
            if info.file_size > limits.max_part_bytes:
                raise UnsafeWorkbookError(f"workbook part exceeds size limit: {info.filename}")
            total += info.file_size
            if total > limits.max_uncompressed_bytes:
                raise UnsafeWorkbookError("workbook exceeds the uncompressed-size limit")
            ratio = info.file_size / max(1, info.compress_size)
            if ratio > limits.max_compression_ratio:
                raise UnsafeWorkbookError(
                    f"workbook part exceeds compression-ratio limit: {info.filename}"
                )
            if _needs_text_scan(info.filename, info.file_size):
                data = archive.read(info)
                _validate_text_part(info.filename, data, limits)


def delimit_workbook_text(text: str, *, source: str, max_characters: int = 100_000) -> str:
    """Place workbook text in a non-instruction data envelope."""

    if len(text) > max_characters:
        raise UnsafeWorkbookError("workbook-derived text exceeds prompt data limit")
    escaped_source = source.replace('"', "'").replace("<", "[").replace(">", "]")
    return (
        f'<UNTRUSTED_WORKBOOK_DATA source="{escaped_source}">\n{text}\n</UNTRUSTED_WORKBOOK_DATA>'
    )


def _validate_member_name(name: str) -> None:
    normalized = PurePosixPath(name.replace("\\", "/"))
    if normalized.is_absolute() or ".." in normalized.parts or not normalized.parts:
        raise UnsafeWorkbookError(f"unsafe workbook package path: {name}")
    if normalized.parts[0].endswith(":"):
        raise UnsafeWorkbookError(f"unsafe workbook package path: {name}")


def _needs_text_scan(name: str, size: int) -> bool:
    lowered = name.casefold()
    return size <= 4 * 1024**2 and (
        lowered.endswith((".xml", ".rels", *_MACRO_SUFFIXES))
        or "vba" in lowered
        or "macro" in lowered
    )


def _validate_text_part(name: str, data: bytes, limits: WorkbookInputLimits) -> None:
    lowered = name.casefold()
    if lowered.endswith(_MACRO_SUFFIXES) and len(data.decode("utf-8", errors="replace")) > (
        limits.max_macro_characters
    ):
        raise UnsafeWorkbookError(f"macro source exceeds size limit: {name}")
    for match in _FORMULA_PATTERN.finditer(data):
        if len(match.group(1).decode("utf-8", errors="replace")) > limits.max_formula_characters:
            raise UnsafeWorkbookError(f"formula exceeds size limit in workbook part: {name}")
