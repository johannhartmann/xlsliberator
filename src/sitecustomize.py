"""Fail closed before host Python can import LibreOffice's UNO bridge.

Python imports ``sitecustomize`` during interpreter startup when this source
tree (or the installed module) is on ``sys.path``.  The office worker is the
only exception, and it must prove both the container marker and the pinned
LibreOffice Python executable path.
"""

from __future__ import annotations

import importlib.abc
import os
import sys
from pathlib import Path

_BLOCKED_MODULES = frozenset({"pyuno", "uno", "unohelper"})
_OFFICE_CONTAINER_MARKER = "XLSLIBERATOR_OFFICE_CONTAINER"
_OFFICE_PYTHON_PREFIX = "/opt/libreoffice26.2/program/"
_SOURCE_RUNTIME_PREFIX = "/opt/libreoffice/program"


def _authorized_python_prefix() -> str:
    candidate = os.environ.get("XLSLIBERATOR_OFFICE_PYTHON_PREFIX", _OFFICE_PYTHON_PREFIX)
    if candidate == _OFFICE_PYTHON_PREFIX:
        return candidate
    resolved = str(Path(candidate).resolve())
    if (
        os.environ.get("XLSLIBERATOR_SOURCE_BUILD_CONTAINER") == "1"
        and resolved.startswith("/office-work/worktrees/")
        and resolved.endswith("/instdir/program")
    ):
        return f"{resolved}/"
    if (
        os.environ.get("XLSLIBERATOR_SOURCE_RUNTIME_CONTAINER") == "1"
        and resolved == _SOURCE_RUNTIME_PREFIX
    ):
        return f"{resolved}/"
    return _OFFICE_PYTHON_PREFIX


def _office_runtime_is_authorized() -> bool:
    executable = str(Path(sys.executable).resolve())
    return (
        os.environ.get(_OFFICE_CONTAINER_MARKER) == "1"
        and Path("/.dockerenv").is_file()
        and executable.startswith(_authorized_python_prefix())
    )


class _HostUnoImportBlocker(importlib.abc.MetaPathFinder):
    """Reject direct and indirect host imports before module discovery."""

    def find_spec(
        self,
        fullname: str,
        path: object = None,
        target: object = None,
    ) -> None:
        del path, target
        if fullname.partition(".")[0] in _BLOCKED_MODULES and not _office_runtime_is_authorized():
            raise ImportError(
                f"Importing {fullname!r} is forbidden outside the pinned LibreOffice Docker runtime"
            )
        return None


if not any(isinstance(finder, _HostUnoImportBlocker) for finder in sys.meta_path):
    sys.meta_path.insert(0, _HostUnoImportBlocker())
