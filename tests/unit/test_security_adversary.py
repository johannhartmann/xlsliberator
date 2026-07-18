"""Security capability and adversary-evaluation tests."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from xlsliberator.scenarios.models import (
    EnvironmentManifest,
    ExternalCapability,
    ExternalCapabilityKind,
)
from xlsliberator.security_adversary import (
    SecurityAdversaryResult,
    SecurityProbe,
    SecurityThreat,
    evaluate_security_probes,
)


def _probe(threat: SecurityThreat, status: str = "BLOCKED") -> SecurityProbe:
    return SecurityProbe.model_validate(
        {
            "threat": threat,
            "status": status,
            "evidence_path": f"evidence/security/{threat.value}.json",
            "detail": "attack was blocked by the declared sandbox boundary",
        }
    )


def test_explicit_open_capability_grants_are_typed_and_recordable() -> None:
    capabilities = [
        ExternalCapability(
            capability=f"migration.{kind.value}",
            kind=kind,
            resource=f"adapter:{kind.value}",
            granted=True,
            constraints={"grant_id": f"grant-{kind.value}"},
        )
        for kind in (
            ExternalCapabilityKind.MAIL,
            ExternalCapabilityKind.DATABASE,
            ExternalCapabilityKind.HTTP,
            ExternalCapabilityKind.FILESYSTEM_EXPORT,
            ExternalCapabilityKind.BUILD_FARM,
        )
    ]

    manifest = EnvironmentManifest(typed_capabilities=capabilities)

    assert manifest.all_granted_capabilities == {
        "migration.mail",
        "migration.database",
        "migration.http",
        "migration.filesystem_export",
        "migration.build_farm",
    }


def test_security_adversary_passes_only_with_every_blocked_probe() -> None:
    result = evaluate_security_probes([_probe(threat) for threat in SecurityThreat])

    assert result.verdict == "PASS"
    assert len(result.probes) == 12


def test_escape_beats_unavailable_and_missing_probe_cannot_pass() -> None:
    probes = [_probe(threat) for threat in SecurityThreat]
    probes[0] = _probe(SecurityThreat.HOST_FILE_ACCESS, "UNAVAILABLE")
    probes[1] = _probe(SecurityThreat.PATH_OR_SYMLINK_ESCAPE, "ESCAPED")

    assert evaluate_security_probes(probes).verdict == "FAIL"

    with pytest.raises(ValidationError, match="at least 12"):
        SecurityAdversaryResult(probes=probes[:-1], verdict="PASS")
