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


def test_ci_pytest_runs_do_not_write_cache_or_basetemp_to_checkout() -> None:
    compose = (ROOT / "docker-compose.yml").read_text(encoding="utf-8")
    ci_check = (ROOT / "tools" / "ci_check.py").read_text(encoding="utf-8")

    assert "PYTEST_ADDOPTS: --basetemp=/tmp/pytest-tmp" in compose
    assert "artifacts/pytest-tmp" not in compose
    assert ci_check.count('"no:cacheprovider"') == 3


def test_package_job_provisions_writable_attestation_directory() -> None:
    workflow = (ROOT / ".github" / "workflows" / "ci.yml").read_text(encoding="utf-8")
    package_job = workflow.split("\n  package:\n", maxsplit=1)[1].split(
        "\n  security:\n", maxsplit=1
    )[0]

    assert "mkdir -p artifacts/ci dist" in package_job
    assert "chmod -R 0777 artifacts dist" in package_job


def test_make_all_includes_runtime_web_and_package_gates() -> None:
    makefile = (ROOT / "Makefile").read_text(encoding="utf-8")

    all_rule = next(line for line in makefile.splitlines() if line.startswith("all:"))
    assert "test-integration" in all_rule
    assert "test-docker-web" in all_rule
    assert "test-package" in all_rule
