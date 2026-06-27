from typing import Any

from click.testing import CliRunner

from xlsliberator.cli import cli


def test_cli_help_includes_web_serve() -> None:
    result = CliRunner().invoke(cli, ["--help"])

    assert result.exit_code == 0
    assert "web-serve" in result.output


def test_web_serve_invokes_uvicorn(monkeypatch: Any) -> None:
    calls: dict[str, Any] = {}

    class FakeUvicorn:
        @staticmethod
        def run(app: str, **kwargs: Any) -> None:
            calls["app"] = app
            calls.update(kwargs)

    monkeypatch.setitem(__import__("sys").modules, "uvicorn", FakeUvicorn)

    result = CliRunner().invoke(
        cli,
        ["web-serve", "--host", "127.0.0.1", "--port", "9001", "--reload"],
    )

    assert result.exit_code == 0
    assert calls == {
        "app": "xlsliberator.web.app:create_app",
        "host": "127.0.0.1",
        "port": 9001,
        "reload": True,
        "factory": True,
    }
