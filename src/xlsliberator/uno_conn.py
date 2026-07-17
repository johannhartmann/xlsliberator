"""Retired host-UNO compatibility surface.

UNO is deliberately available only inside :mod:`xlsliberator.lo_worker`, which
runs in the pinned LibreOffice container.  Keeping these names lets older callers
receive an explicit failure without ever importing PyUNO or starting office on
the host.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, NoReturn


class UnoConnectionError(RuntimeError):
    """Raised whenever legacy direct UNO access is requested."""


_DOCKER_ONLY_MESSAGE = (
    "Direct UNO access is disabled; use LibreOfficeWorkerClient so PyUNO and "
    "LibreOffice execute only inside the pinned Docker runtime"
)


def _unsupported() -> NoReturn:
    raise UnoConnectionError(_DOCKER_ONLY_MESSAGE)


class UnoCtx:
    """Fail-closed compatibility shim for the removed host UNO context."""

    def __init__(self, *_args: Any, **_kwargs: Any) -> None:
        self.desktop: Any = None
        self.component_context: Any = None

    def __enter__(self) -> UnoCtx:
        _unsupported()

    def __exit__(self, _exc_type: Any, _exc_val: Any, _exc_tb: Any) -> None:
        return None

    def connect(self) -> None:
        _unsupported()

    def disconnect(self) -> None:
        return None

    @property
    def is_connected(self) -> bool:
        return False


def connect_lo(*_args: Any, **_kwargs: Any) -> UnoCtx:
    _unsupported()


def new_calc(_ctx: UnoCtx) -> Any:
    _unsupported()


def set_macro_security_level(_ctx: UnoCtx, level: int = 0) -> None:
    del level
    _unsupported()


def open_calc(_ctx: UnoCtx, path: str | Path) -> Any:
    del path
    _unsupported()


def save_as_ods(_ctx: UnoCtx, _doc: Any, path: str | Path) -> None:
    del path
    _unsupported()


def recalc(_ctx: UnoCtx, _doc: Any) -> None:
    _unsupported()


def get_sheet(_ctx: UnoCtx, _doc: Any, name_or_index: str | int) -> Any:
    del name_or_index
    _unsupported()


def get_cell(_ctx: UnoCtx, _sheet: Any, address: str) -> Any:
    del address
    _unsupported()
