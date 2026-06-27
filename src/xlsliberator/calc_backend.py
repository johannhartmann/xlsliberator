"""Runtime backend discovery and isolated office profiles."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

from xlsliberator.validation_models import TargetKind

VERSION_TIMEOUT_SECONDS = 5
UNO_START_TIMEOUT_SECONDS = 20
UNO_PARSE_TIMEOUT_SECONDS = 30


@dataclass(frozen=True)
class CalcBackendInfo:
    """Discovered office backend information."""

    kind: TargetKind
    executable: str
    version: str | None = None
    available: bool = False


@dataclass
class RuntimeProfile:
    """Isolated office runtime profile."""

    user_installation_dir: Path
    env: dict[str, str] = field(default_factory=dict)
    cleanup: bool = True

    @property
    def user_installation_url(self) -> str:
        """Return the file URL expected by -env:UserInstallation."""
        return self.user_installation_dir.resolve().as_uri()

    @property
    def user_installation_arg(self) -> str:
        """Return the command-line argument for this isolated profile."""
        return f"-env:UserInstallation={self.user_installation_url}"


class CalcBackend(Protocol):
    """Protocol implemented by office runtime backends."""

    info: CalcBackendInfo

    def version(self) -> str | None:
        """Return the backend version string if available."""

    def parse_formula_text(self, formula: str, sheet_name: str | None = None) -> object:
        """Parse target formula text when parser integration is available."""


class _BaseBackend:
    """Base implementation shared by concrete backends."""

    kind: TargetKind

    def __init__(self, executable: str, version: str | None = None) -> None:
        self.info = CalcBackendInfo(
            kind=self.kind,
            executable=executable,
            version=version,
            available=True,
        )

    def version(self) -> str | None:
        """Return the backend version string."""
        return self.info.version

    def parse_formula_text(self, formula: str, sheet_name: str | None = None) -> object:
        """Validate formula text with UNO FormulaParser when available."""
        from xlsliberator.formula_engine import FormulaDialect, FormulaEngine

        uno_result = _parse_with_uno_formula_parser(
            executable=self.info.executable,
            formula=formula,
            sheet_name=sheet_name,
            backend_kind=self.info.kind.value,
            backend_version=self.info.version,
        )
        if uno_result is not None:
            return uno_result

        result = FormulaEngine().validate_formula_text(formula, FormulaDialect.CALC_A1)
        result.details.update(
            {
                "backend_kind": self.info.kind.value,
                "backend_version": self.info.version,
                "sheet_name": sheet_name,
                "target_parser": "basic_structural_fallback",
                "target_parser_unavailable": "UNO worker unavailable",
            }
        )
        return result


class LibreOfficeBackend(_BaseBackend):
    """LibreOffice runtime backend."""

    kind = TargetKind.LIBREOFFICE


class ApacheOpenOfficeBackend(_BaseBackend):
    """Apache OpenOffice runtime backend."""

    kind = TargetKind.OPENOFFICE


def detect_version(executable: str) -> str | None:
    """Detect a backend version string with a short timeout."""
    try:
        result = subprocess.run(
            [executable, "--version"],
            capture_output=True,
            text=True,
            timeout=VERSION_TIMEOUT_SECONDS,
            check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return None

    output = (result.stdout or result.stderr).strip()
    return output or None


def discover_backends() -> list[CalcBackend]:
    """Discover available LibreOffice and Apache OpenOffice backends."""
    backends: list[CalcBackend] = []
    claimed_executables: set[str] = set()

    for executable_name in ("soffice", "libreoffice"):
        executable = shutil.which(executable_name)
        if executable is None or executable in claimed_executables:
            continue
        version = detect_version(executable)
        # Skip only when the binary positively identifies as OpenOffice.
        if version and "openoffice" in version.lower() and "libreoffice" not in version.lower():
            continue
        backends.append(LibreOfficeBackend(executable, version))
        claimed_executables.add(executable)

    for executable_name in ("openoffice", "soffice.bin", "soffice"):
        executable = shutil.which(executable_name)
        if executable is None or executable in claimed_executables:
            continue
        version = detect_version(executable)
        version_text = (version or "").lower()
        if "libreoffice" in version_text:
            continue
        # `soffice`/`soffice.bin` are shared with LibreOffice, so an unknown
        # version must not be assumed to be OpenOffice (that would mislabel a
        # LibreOffice-only host). Only the unambiguous `openoffice` binary, or a
        # version string that names OpenOffice, qualifies.
        if executable_name in ("soffice", "soffice.bin") and "openoffice" not in version_text:
            continue
        backends.append(ApacheOpenOfficeBackend(executable, version))
        claimed_executables.add(executable)

    return backends


def _parse_with_uno_formula_parser(
    *,
    executable: str,
    formula: str,
    sheet_name: str | None,
    backend_kind: str,
    backend_version: str | None,
) -> Any | None:
    """Parse a Calc formula through the out-of-process UNO FormulaParser worker."""
    from xlsliberator.formula_engine import FormulaDialect, FormulaParseResult
    from xlsliberator.lo_worker_client import (
        UNO_WORKER_UNAVAILABLE,
        LibreOfficeWorkerClient,
    )

    client = LibreOfficeWorkerClient(
        office_executable=executable,
        timeout_seconds=UNO_PARSE_TIMEOUT_SECONDS,
    )
    ping = client.ping()
    if not ping.success:
        return None

    response = client.request(
        {
            "op": "parse_formula",
            "formula": formula,
            "sheet_name": sheet_name,
            "timeout_seconds": UNO_PARSE_TIMEOUT_SECONDS,
        }
    )
    if response.success:
        return FormulaParseResult(
            success=True,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            tokens=list(response.data.get("tokens") or []),
            details={
                "backend_kind": backend_kind,
                "backend_version": backend_version,
                "sheet_name": sheet_name,
                "target_parser": "uno_formula_parser",
                "roundtrip_formula": response.data.get("roundtrip_formula"),
                "worker_wrapper": response.wrapper_path,
            },
        )

    if response.error and response.error.type in {
        "UnoWorkerUnavailable",
        "ImportError",
        "MalformedWorkerJSON",
        "TimeoutExpired",
    }:
        return FormulaParseResult(
            success=False,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            error=response.error.message,
            details={
                "backend_kind": backend_kind,
                "backend_version": backend_version,
                "sheet_name": sheet_name,
                "target_parser": "uno_formula_parser",
                "target_parser_unavailable": UNO_WORKER_UNAVAILABLE,
                "worker_error_type": response.error.type,
                "worker_wrapper": response.wrapper_path,
            },
        )

    return FormulaParseResult(
        success=False,
        formula=formula,
        dialect=FormulaDialect.CALC_A1,
        error=response.error.message if response.error else "UNO worker request failed",
        details={
            "backend_kind": backend_kind,
            "backend_version": backend_version,
            "sheet_name": sheet_name,
            "target_parser": "uno_formula_parser",
            "parser_exception": response.error.type if response.error else "WorkerError",
            "worker_wrapper": response.wrapper_path,
        },
    )


@contextmanager
def create_isolated_user_profile(prefix: str) -> Iterator[RuntimeProfile]:
    """Create an isolated office user profile for runtime validation."""
    with tempfile.TemporaryDirectory(prefix=prefix) as tmpdir:
        profile_dir = Path(tmpdir) / "user"
        profile_dir.mkdir(parents=True, exist_ok=True)
        env = dict(os.environ)
        yield RuntimeProfile(user_installation_dir=profile_dir, env=env, cleanup=True)
