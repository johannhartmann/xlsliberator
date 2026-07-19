"""Private filesystem state shared by the Open-SWE graph and its HTTP adapter."""

from __future__ import annotations

import json
import os
import threading
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

_STATE_LOCK = threading.RLock()


def workspace_root() -> Path:
    """Return the container-only root containing isolated migration workspaces."""
    return Path(os.environ.get("XLSLIBERATOR_OPEN_SWE_WORKSPACE_ROOT", "/workspaces"))


def thread_root(thread_id: str) -> Path:
    """Return one canonical thread workspace without accepting path fragments."""
    if not thread_id or any(character not in "0123456789abcdef-" for character in thread_id):
        raise ValueError("invalid migration thread identifier")
    return workspace_root() / thread_id


def state_path(thread_id: str) -> Path:
    return thread_root(thread_id) / "migration.json"


def read_state(thread_id: str) -> dict[str, Any]:
    """Read a migration record."""
    payload = json.loads(state_path(thread_id).read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise ValueError("migration state is not an object")
    return payload


def write_state(thread_id: str, payload: dict[str, Any]) -> None:
    """Atomically replace a migration record."""
    with _STATE_LOCK:
        destination = state_path(thread_id)
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.with_suffix(".json.tmp")
        temporary.write_text(
            json.dumps(payload, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        temporary.replace(destination)


def update_state(thread_id: str, **updates: Any) -> dict[str, Any]:
    """Apply an atomic shallow update to a migration record."""
    with _STATE_LOCK:
        payload = read_state(thread_id)
        payload.update(updates)
        payload["updated_at"] = _timestamp()
        write_state(thread_id, payload)
        return payload


def append_event(
    thread_id: str,
    *,
    stage: str,
    message: str,
    status: str = "running",
) -> dict[str, Any]:
    """Append a public progress event without model reasoning or tool arguments."""
    with _STATE_LOCK:
        payload = read_state(thread_id)
        events = payload.setdefault("events", [])
        if not isinstance(events, list):
            raise ValueError("migration events are not a list")
        event = {
            "index": len(events),
            "stage": stage,
            "message": message[:500],
            "status": status,
            "timestamp": _timestamp(),
        }
        events.append(event)
        payload["updated_at"] = event["timestamp"]
        write_state(thread_id, payload)
        return event


def _timestamp() -> str:
    return datetime.now(UTC).isoformat()
