"""Explicitly authorized live-model run against the embedded Open-SWE service."""

from __future__ import annotations

import os
import time
from pathlib import Path

import openpyxl
import pytest

from xlsliberator.web.open_swe import OpenSWEClient


@pytest.mark.integration
@pytest.mark.live
def test_real_local_langgraph_workbook_migration_when_available(tmp_path: Path) -> None:
    if os.getenv("XLSLIBERATOR_ALLOW_PAID_OPEN_SWE_TEST") != "1":
        pytest.skip("Set XLSLIBERATOR_ALLOW_PAID_OPEN_SWE_TEST=1 to authorize provider usage")
    base_url = "http://xlsliberator-open-swe:2024"
    token = os.getenv("XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN")
    owner_id = os.getenv("XLSLIBERATOR_OPEN_SWE_OWNER_ID", "docker-real-smoke")
    if not token:
        pytest.fail("XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN is required")

    source = tmp_path / "real-langgraph-smoke.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "Acceptance"
    workbook.active["A1"] = "input"
    workbook.active["B1"] = "result"
    workbook.active["A2"] = 2
    workbook.active["B2"] = "=A2*3"
    workbook.save(source)
    workbook.close()

    client = OpenSWEClient(
        base_url=base_url,
        token=token,
        owner_id=owner_id,
        timeout_seconds=60,
    )
    created = client.create_migration(
        source,
        requirements="Preserve the formula behavior after save, close, and reopen.",
    )
    thread_id = str(created["thread_id"])
    try:
        deadline = time.monotonic() + int(
            os.getenv("XLSLIBERATOR_OPEN_SWE_JOB_TIMEOUT_SECONDS", "1800")
        )
        status: dict[str, object] = {}
        while time.monotonic() < deadline:
            status = client.status(thread_id)
            if status.get("status") in {"complete", "failed", "cancelled", "rejected"}:
                break
            time.sleep(2)

        assert status.get("status") == "complete", status
        artifacts = status.get("artifacts")
        assert isinstance(artifacts, list)
        names = {artifact.get("name") for artifact in artifacts if isinstance(artifact, dict)}
        assert {"target.ods", "report.json", "save-reopen.json"}.issubset(names)
    finally:
        client.cleanup(thread_id)
