"""Static contracts for required CI gates."""

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def test_required_integration_jobs_are_blocking_and_upload_failure_evidence() -> None:
    workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")

    assert "office-integration:" in workflow
    assert "docker-web:" in workflow
    assert "continue-on-error:" not in workflow
    assert "|| true" not in workflow
    assert workflow.count("if: always()") >= 3
    assert "office-integration-evidence" in workflow
    assert "name: docker-web" in workflow
    assert "set -o pipefail" in workflow
    assert "artifacts/ci/office-integration.log" in workflow
    assert "artifacts/ci/docker-web.log" in workflow


def test_make_all_includes_runtime_web_and_package_gates() -> None:
    makefile = (ROOT / "Makefile").read_text(encoding="utf-8")

    all_rule = next(line for line in makefile.splitlines() if line.startswith("all:"))
    assert "test-integration" in all_rule
    assert "test-docker-web" in all_rule
    assert "test-package" in all_rule
