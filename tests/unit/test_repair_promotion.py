"""Tests for reusable repair promotion and its MCP service boundaries."""

from __future__ import annotations

import asyncio
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from click.testing import CliRunner
from pydantic import ValidationError

from xlsliberator.cli import cli
from xlsliberator.migration_services_mcp import (
    _serve,
    collect_build_logs,
    compare_stock_patched,
    create_source_worktree,
    register_minimized_failure,
    run_hidden_acceptance,
    run_public_suite,
    search_prior_failures,
    search_public_fixtures,
)
from xlsliberator.repair_promotion import RepairRecord, load_repair_records

ROOT = Path(__file__).parents[2]
RECORD_PATH = ROOT / "repairs/tdf-172479-text-functions/record.json"


def test_real_libreoffice_repair_record_is_complete_and_verified() -> None:
    record = RepairRecord.load(RECORD_PATH)

    assert record.classification == "libreoffice"
    assert record.libreoffice is not None
    assert record.libreoffice.full_build == "26.2.4.2"
    assert record.upstream_review == "https://gerrit.libreoffice.org/c/core/+/206776"
    assert not record.verify(ROOT)
    assert load_repair_records(ROOT / "repairs") == [record]


def test_repair_record_rejects_layer_switch_and_missing_stage() -> None:
    payload = json.loads(RECORD_PATH.read_text(encoding="utf-8"))
    payload["fixed_layer"] = "test-validation"
    with pytest.raises(ValidationError, match="classification and patched layer"):
        RepairRecord.model_validate(payload)

    payload["fixed_layer"] = "libreoffice"
    payload["stages"] = payload["stages"][:-1]
    with pytest.raises(ValidationError):
        RepairRecord.model_validate(payload)


def test_public_corpus_tools_search_and_validate_real_repair() -> None:
    fixtures = asyncio.run(search_public_fixtures("text formula"))
    failures = asyncio.run(search_prior_failures("textafter libreoffice"))
    suite = asyncio.run(run_public_suite("tdf-172479-text-functions"))
    comparison = asyncio.run(compare_stock_patched("tdf-172479-text-functions"))
    logs = asyncio.run(collect_build_logs("tdf-172479-text-functions"))

    assert fixtures["success"] is True
    assert fixtures["hidden_expectations_included"] is False
    assert any(item["fixture_id"] == "public-tdf-172479" for item in fixtures["fixtures"])
    assert failures["matches"][0]["repair_id"] == "tdf-172479-text-functions"
    assert suite["success"] is True
    assert suite["stock_disposition"] == "failed-as-expected"
    assert suite["patched_disposition"] == "passed"
    assert comparison["success"] is True
    assert logs["success"] is True
    assert logs["runtime_identity"]["full_build"] == "26.2.4.2"


def test_corpus_registration_is_fail_closed_and_records_not_executed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    denied = asyncio.run(
        register_minimized_failure(
            "new-regression",
            "stable-signature",
            {"cells": ["A1"]},
            "generated",
            "CC0-1.0",
        )
    )
    assert denied["success"] is False
    assert denied["error"]["type"] == "CorpusRegistrationUnauthorized"

    monkeypatch.setenv("XLSLIBERATOR_CORPUS_REGISTRATION_ENABLED", "true")
    monkeypatch.setenv("XLSLIBERATOR_CORPUS_REGISTRY_ROOT", str(tmp_path))
    registered = asyncio.run(
        register_minimized_failure(
            "new-regression",
            "stable-signature",
            {"cells": ["A1"]},
            "generated",
            "CC0-1.0",
        )
    )

    assert registered["success"] is True
    payload = json.loads((tmp_path / "new-regression.json").read_text(encoding="utf-8"))
    assert payload["status"] == "registered-not-executed"


def test_hidden_corpus_and_buildfarm_mutation_require_server_authority(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    hidden = asyncio.run(run_hidden_acceptance("tdf-172479-text-functions"))
    mutation = asyncio.run(
        create_source_worktree(
            "0229ac93fcf0d7cbc6376066c6f35021cef002dc",
            "tdf-172479-text-functions",
        )
    )
    assert hidden["success"] is False
    assert hidden["error"]["type"] == "HiddenCorpusUnauthorized"
    assert mutation["success"] is False
    assert mutation["error"]["type"] == "BuildFarmUnauthorized"

    monkeypatch.setenv("XLSLIBERATOR_BUILD_FARM_MUTATION_ENABLED", "true")
    unavailable = asyncio.run(
        create_source_worktree(
            "0229ac93fcf0d7cbc6376066c6f35021cef002dc",
            "tdf-172479-text-functions",
        )
    )
    assert unavailable["operation_status"] == "unavailable"
    assert unavailable["capability_available"] is False


def test_migration_mcp_rejects_public_bind_without_trusted_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    server = MagicMock()
    monkeypatch.setenv("XLSLIBERATOR_APPLICATION_CONTAINER", "1")
    monkeypatch.delenv("XLSLIBERATOR_MCP_TRUSTED_CONTAINER_PROXY", raising=False)

    with pytest.raises(ValueError, match="trusted-container proxy"):
        _serve(server, "0.0.0.0", 8010)

    server.run.assert_not_called()


def test_migration_mcp_allows_explicit_trusted_container_proxy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    server = MagicMock()
    monkeypatch.setenv("XLSLIBERATOR_APPLICATION_CONTAINER", "1")
    monkeypatch.setenv("XLSLIBERATOR_MCP_TRUSTED_CONTAINER_PROXY", "1")

    _serve(server, "0.0.0.0", 8010)

    server.run.assert_called_once_with(transport="http", host="0.0.0.0", port=8010)


def test_repair_service_cli_commands_are_registered() -> None:
    runner = CliRunner()

    corpus = runner.invoke(cli, ["corpus-mcp-serve", "--help"])
    buildfarm = runner.invoke(cli, ["buildfarm-mcp-serve", "--help"])

    assert corpus.exit_code == 0
    assert buildfarm.exit_code == 0
