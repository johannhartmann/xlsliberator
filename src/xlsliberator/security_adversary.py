"""Deterministic security-adversary evaluation for untrusted workbook execution."""

from __future__ import annotations

from enum import StrEnum
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator


class SecurityThreat(StrEnum):
    HOST_FILE_ACCESS = "host_file_access"
    PATH_OR_SYMLINK_ESCAPE = "path_or_symlink_escape"
    NETWORK_EXFILTRATION = "network_exfiltration"
    PROCESS_PERSISTENCE = "process_persistence"
    RESOURCE_EXHAUSTION = "resource_exhaustion"
    ARCHIVE_BOMB = "archive_bomb"
    MALFORMED_DOCUMENT = "malformed_document"
    MACRO_NONTERMINATION = "macro_nontermination"
    PROMPT_INJECTION = "prompt_injection"
    UNAUTHORIZED_MCP = "unauthorized_mcp"
    HIDDEN_TEST_LEAKAGE = "hidden_test_leakage"
    CROSS_JOB_ACCESS = "cross_job_access"


class SecurityProbe(BaseModel):
    """One independently evaluated attack probe."""

    model_config = ConfigDict(extra="forbid")

    threat: SecurityThreat
    status: Literal["BLOCKED", "ESCAPED", "UNAVAILABLE"]
    evidence_path: str = Field(min_length=1, max_length=500)
    detail: str = Field(min_length=1, max_length=1000)


class SecurityAdversaryResult(BaseModel):
    """Fail-closed aggregate; narratives cannot replace missing probes."""

    model_config = ConfigDict(extra="forbid")

    schema_version: Literal["1.0.0"] = "1.0.0"
    evaluator: Literal["security-adversary-v1"] = "security-adversary-v1"
    target_libreoffice_build: Literal["26.2.4.2"] = "26.2.4.2"
    probes: list[SecurityProbe] = Field(min_length=12, max_length=12)
    verdict: Literal["PASS", "FAIL", "UNAVAILABLE"]

    @model_validator(mode="after")
    def verdict_matches_complete_probe_set(self) -> SecurityAdversaryResult:
        expected = set(SecurityThreat)
        actual = {probe.threat for probe in self.probes}
        if actual != expected or len(actual) != len(self.probes):
            raise ValueError("security evaluation requires every threat exactly once")
        statuses = {probe.status for probe in self.probes}
        required = (
            "FAIL"
            if "ESCAPED" in statuses
            else "UNAVAILABLE"
            if "UNAVAILABLE" in statuses
            else "PASS"
        )
        if self.verdict != required:
            raise ValueError(f"security verdict must be {required}")
        return self


def evaluate_security_probes(probes: list[SecurityProbe]) -> SecurityAdversaryResult:
    """Derive the only truthful aggregate verdict from independent probe results."""
    statuses = {probe.status for probe in probes}
    verdict: Literal["PASS", "FAIL", "UNAVAILABLE"]
    if "ESCAPED" in statuses:
        verdict = "FAIL"
    elif "UNAVAILABLE" in statuses:
        verdict = "UNAVAILABLE"
    else:
        verdict = "PASS"
    return SecurityAdversaryResult(probes=probes, verdict=verdict)
