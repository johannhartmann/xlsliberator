"""Tests for Docker Calc backend discovery and runtime profiles."""

from pathlib import Path
from typing import Any

from xlsliberator import calc_backend
from xlsliberator.calc_backend import create_isolated_user_profile, discover_backends
from xlsliberator.docker_runtime import DockerRuntimeIdentity, DockerRuntimeUnavailable
from xlsliberator.validation_models import TargetKind


class FakeRuntime:
    def __init__(
        self,
        identity: DockerRuntimeIdentity | None = None,
        unavailable: bool = False,
        response: dict[str, Any] | None = None,
    ):
        self.identity = identity
        self.unavailable = unavailable
        self.response = response

    def resolve_identity(self) -> DockerRuntimeIdentity:
        if self.unavailable:
            raise DockerRuntimeUnavailable("missing")
        assert self.identity is not None
        return self.identity

    def request(self, payload: dict[str, Any], *, _identity: str | None = None) -> dict[str, Any]:
        assert payload["op"] == "parse_formula"
        assert _identity == "sha256:abc"
        if self.unavailable:
            raise DockerRuntimeUnavailable("missing")
        assert self.response is not None
        return self.response


def test_discover_backends_when_docker_runtime_missing() -> None:
    assert discover_backends(FakeRuntime(unavailable=True)) == []  # type: ignore[arg-type]


def test_discover_backends_uses_immutable_docker_identity() -> None:
    identity = DockerRuntimeIdentity(
        image_reference="xlsliberator-libreoffice:26.2.4.2",
        image_id="sha256:abc",
        version="26.2.4.2",
    )

    backends = discover_backends(FakeRuntime(identity))  # type: ignore[arg-type]

    assert len(backends) == 1
    assert backends[0].info.kind == TargetKind.LIBREOFFICE
    assert backends[0].info.executable == "sha256:abc"
    assert backends[0].info.version == "26.2.4.2"


def test_create_isolated_user_profile_url() -> None:
    with create_isolated_user_profile("xlsliberator-test-") as profile:
        assert profile.user_installation_dir.exists()
        assert profile.user_installation_url.startswith("file://")
        assert profile.user_installation_arg.startswith("-env:UserInstallation=file://")
        assert isinstance(profile.env, dict)

    assert not Path(profile.user_installation_dir).exists()


def test_backend_formula_parse_hook_uses_uno_when_available(monkeypatch: Any) -> None:
    del monkeypatch
    runtime = FakeRuntime(
        response={
            "success": True,
            "data": {
                "tokens": ["token"],
                "roundtrip_formula": "=SUM(1;2)",
                "roundtrip_equivalent": True,
                "parser_accepted": True,
                "syntax_errors": [],
                "container_image_id": "sha256:abc",
            },
        }
    )
    backend = calc_backend.LibreOfficeBackend(
        "sha256:abc",
        "26.2.4.2",
        runtime=runtime,  # type: ignore[arg-type]
    )

    result = backend.parse_formula_text("=SUM(1;2)", sheet_name="Sheet1")

    assert result.success
    assert result.details["target_parser"] == "docker_uno_formula_parser"


def test_backend_formula_parse_hook_reports_structural_fallback(monkeypatch: Any) -> None:
    del monkeypatch
    backend = calc_backend.LibreOfficeBackend("sha256:abc", "26.2.4.2")
    backend.runtime = FakeRuntime(unavailable=True)  # type: ignore[assignment]

    result = backend.parse_formula_text("=SUM(1;2)", sheet_name="Sheet1")

    assert not result.success
    assert result.details["backend_kind"] == "libreoffice"
    assert result.details["target_parser"] == "docker_uno_formula_parser"
    assert "target_parser_unavailable" in result.details
