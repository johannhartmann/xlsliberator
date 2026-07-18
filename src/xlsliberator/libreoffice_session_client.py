"""Typed client boundary for the session-oriented LibreOffice MCP service."""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from typing import Any, Protocol

from xlsliberator.libreoffice_session import SessionOperation


class SessionTransport(Protocol):
    """Minimal transport used by real MCP adapters and deterministic fakes."""

    async def call(self, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
        """Call one named session tool."""


class LibreOfficeSessionClient:
    """Client that enforces explicit session IDs on every non-create call."""

    def __init__(self, transport: SessionTransport) -> None:
        self.transport = transport

    async def create_session(
        self,
        environment: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        response = await self._call("create_session", {"environment": environment})
        if response.get("success") and not response.get("session_id"):
            raise ValueError("LibreOffice create_session response omitted its session ID")
        return response

    async def call(
        self,
        session_id: str,
        operation: SessionOperation | str,
        **arguments: Any,
    ) -> dict[str, Any]:
        if not session_id:
            raise ValueError("session_id is required")
        tool = operation.value if isinstance(operation, SessionOperation) else operation
        return await self._call(
            tool, {"session_id": session_id, **arguments}, session_id=session_id
        )

    async def _call(
        self,
        tool: str,
        arguments: dict[str, Any],
        *,
        session_id: str | None = None,
    ) -> dict[str, Any]:
        response = await self.transport.call(tool, arguments)
        if not isinstance(response, dict):
            raise TypeError("LibreOffice session transport returned a non-object response")
        if session_id is not None and response.get("session_id") != session_id:
            raise ValueError("LibreOffice session response ID does not match the request")
        required = {"transport_success", "operation_status", "success"}
        missing = required - response.keys()
        if missing:
            raise ValueError(f"LibreOffice session response omitted fields: {sorted(missing)}")
        return response


class InProcessSessionTransport:
    """Deterministic transport adapter for fake server/client tests."""

    def __init__(
        self,
        tools: dict[str, Callable[..., Awaitable[dict[str, Any]]]],
    ) -> None:
        self.tools = dict(tools)
        self.calls: list[tuple[str, dict[str, Any]]] = []

    async def call(self, tool: str, arguments: dict[str, Any]) -> dict[str, Any]:
        self.calls.append((tool, dict(arguments)))
        try:
            function = self.tools[tool]
        except KeyError as exc:
            raise ValueError(f"unknown session tool: {tool}") from exc
        return await function(**arguments)
