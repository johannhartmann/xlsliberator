"""Typed request and response contracts for CLI, MCP, and worker boundaries."""

from __future__ import annotations

import os
from typing import Any, Literal

from pydantic import BaseModel, Field, model_validator

from xlsliberator.docker_runtime import DEFAULT_IMAGE
from xlsliberator.validation_models import GateExecutionStatus


class BoundaryError(BaseModel):
    """Serializable error without conflating it with transport state."""

    type: str
    message: str
    details: dict[str, Any] = Field(default_factory=dict)


class EvidenceRecord(BaseModel):
    """Evidence attached to one boundary operation."""

    kind: str
    data: dict[str, Any] = Field(default_factory=dict)
    path: str | None = None


class RuntimeResourcePolicy(BaseModel):
    """Fixed resource policy exposed in tool schemas but not caller-expandable."""

    network: Literal["none"] = "none"
    read_only_root: Literal[True] = True
    non_root: Literal[True] = True
    cap_drop: tuple[Literal["ALL"], ...] = ("ALL",)
    no_new_privileges: Literal[True] = True
    pids_limit: Literal[256] = 256
    memory: Literal["2g"] = "2g"
    cpus: Literal[2] = 2
    file_size_limit_blocks: Literal[1048576] = 1048576


class RuntimeToolOptions(BaseModel):
    """Caller-visible runtime selection constrained to configured identities."""

    target: Literal["libreoffice"] = "libreoffice"
    target_runtime_image: str = DEFAULT_IMAGE
    target_runtime_digest: str | None = None
    timeout_seconds: int = Field(default=30, ge=1, le=120)
    resource_limits: RuntimeResourcePolicy = Field(default_factory=RuntimeResourcePolicy)
    workspace_root: str | None = None
    evidence_destination: str | None = None

    @model_validator(mode="after")
    def constrain_runtime_identity(self) -> RuntimeToolOptions:
        configured_image = os.environ.get("XLSLIBERATOR_LIBREOFFICE_IMAGE", DEFAULT_IMAGE)
        if self.target_runtime_image != configured_image:
            raise ValueError("target_runtime_image is not a configured runtime identity")
        configured_digest = os.environ.get("XLSLIBERATOR_LIBREOFFICE_IMAGE_ID")
        if self.target_runtime_digest is not None and (
            configured_digest is None or self.target_runtime_digest != configured_digest
        ):
            raise ValueError("target_runtime_digest is not a configured runtime identity")
        return self


class BoundaryResponse(BaseModel):
    """Canonical response shared across process and user-facing boundaries."""

    schema_version: Literal["1.0.0"] = "1.0.0"
    transport_success: bool
    operation_status: GateExecutionStatus
    implemented: bool
    capability_available: bool
    evidence: list[EvidenceRecord] = Field(default_factory=list)
    error: BoundaryError | None = None
    data: dict[str, Any] = Field(default_factory=dict)

    @property
    def success(self) -> bool:
        """Only an actually passed operation is successful."""
        return self.operation_status == GateExecutionStatus.PASSED

    def to_payload(self) -> dict[str, Any]:
        """Return the canonical fields plus backward-compatible flattened data."""
        payload = self.model_dump(mode="json", exclude={"data"})
        payload["success"] = self.success
        payload.update(self.data)
        return payload
