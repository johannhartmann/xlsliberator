"""Runtime backend discovery and isolated office profiles."""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

from xlsliberator.docker_runtime import (
    DockerRuntimeUnavailable,
    LibreOfficeDockerRuntime,
)
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

    def __init__(
        self,
        executable: str,
        version: str | None = None,
        *,
        runtime: LibreOfficeDockerRuntime | None = None,
    ) -> None:
        self.runtime = runtime or LibreOfficeDockerRuntime()
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
        from xlsliberator.formula_engine import FormulaDialect, FormulaEngine, FormulaParseResult

        uno_result = _parse_with_uno_formula_parser(
            runtime=self.runtime,
            image_id=self.info.executable,
            formula=formula,
            sheet_name=sheet_name,
            backend_kind=self.info.kind.value,
            backend_version=self.info.version,
        )
        if uno_result is not None:
            return uno_result

        diagnostic = FormulaEngine().validate_formula_text(formula, FormulaDialect.CALC_A1)
        return FormulaParseResult(
            success=False,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            error="Docker UNO FormulaParser is unavailable; structural diagnostics cannot certify",
            details={
                "backend_kind": self.info.kind.value,
                "backend_version": self.info.version,
                "sheet_name": sheet_name,
                "target_parser": "unavailable",
                "target_parser_unavailable": "Docker worker unavailable",
                "structural_diagnostic": diagnostic.model_dump(mode="json"),
            },
        )


class LibreOfficeBackend(_BaseBackend):
    """LibreOffice runtime backend."""

    kind = TargetKind.LIBREOFFICE


def detect_version(executable: str) -> str | None:
    """Never probe host office executables; Docker identity is authoritative."""
    del executable
    return None


def discover_backends(
    runtime: LibreOfficeDockerRuntime | None = None,
) -> list[CalcBackend]:
    """Discover only the pinned, probed LibreOffice Docker backend."""
    try:
        identity = (runtime or LibreOfficeDockerRuntime()).resolve_identity()
    except DockerRuntimeUnavailable:
        return []
    return [LibreOfficeBackend(identity.image_id, identity.version, runtime=runtime)]


def _parse_with_uno_formula_parser(
    *,
    runtime: LibreOfficeDockerRuntime,
    image_id: str,
    formula: str,
    sheet_name: str | None,
    backend_kind: str,
    backend_version: str | None,
) -> Any | None:
    """Parse a Calc formula only through the disposable Docker UNO worker."""
    from xlsliberator.formula_engine import FormulaDialect, FormulaParseResult

    try:
        response = runtime.request(
            {
                "op": "parse_formula",
                "formula": formula,
                "sheet_name": sheet_name,
                "timeout_seconds": UNO_PARSE_TIMEOUT_SECONDS,
            },
            _identity=image_id,
        )
    except DockerRuntimeUnavailable as exc:
        return FormulaParseResult(
            success=False,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            error=str(exc),
            details={
                "backend_kind": backend_kind,
                "backend_version": backend_version,
                "sheet_name": sheet_name,
                "target_parser": "docker_uno_formula_parser",
                "target_parser_unavailable": str(exc),
            },
        )
    data = dict(response.get("data") or {})
    if response.get("success"):
        parser_accepted = data.get("parser_accepted") is True
        roundtrip_equivalent = data.get("roundtrip_equivalent") is True
        accepted = parser_accepted and roundtrip_equivalent
        syntax_errors = [str(error) for error in data.get("syntax_errors") or []]
        error_message: str | None = None
        if not accepted:
            reasons = syntax_errors.copy()
            if not parser_accepted and not reasons:
                reasons.append("target FormulaParser syntax acceptance was not proven")
            if not roundtrip_equivalent:
                reasons.append("target FormulaParser token round-trip changed")
            error_message = "; ".join(reasons)
        return FormulaParseResult(
            success=accepted,
            formula=formula,
            dialect=FormulaDialect.CALC_A1,
            error=error_message,
            tokens=list(data.get("tokens") or []),
            details={
                "backend_kind": backend_kind,
                "backend_version": backend_version,
                "sheet_name": sheet_name,
                "target_parser": "docker_uno_formula_parser",
                "parser_accepted": parser_accepted,
                "syntax_errors": syntax_errors,
                "roundtrip_equivalent": roundtrip_equivalent,
                "roundtrip_formula": data.get("roundtrip_formula"),
                "container_image_id": data.get("container_image_id"),
                "container_name": data.get("container_name"),
            },
        )
    worker_error = dict(response.get("error") or {})
    return FormulaParseResult(
        success=False,
        formula=formula,
        dialect=FormulaDialect.CALC_A1,
        error=str(worker_error.get("message") or "Docker UNO worker request failed"),
        details={
            "backend_kind": backend_kind,
            "backend_version": backend_version,
            "sheet_name": sheet_name,
            "target_parser": "docker_uno_formula_parser",
            "parser_exception": worker_error.get("type") or "WorkerError",
            "container_image_id": data.get("container_image_id"),
            "container_name": data.get("container_name"),
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
