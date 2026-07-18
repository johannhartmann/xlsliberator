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
    docker_web_job = workflow.split("\n  docker-web:\n", maxsplit=1)[1].split(
        "\n  package:\n", maxsplit=1
    )[0]
    assert "sudo rm -rf artifacts/ci/docker-web-tmp" in docker_web_job
    assert "path: artifacts/" not in docker_web_job
    assert "artifacts/ci/docker-web-build.log" in docker_web_job
    assert "artifacts/ci/docker-web-attestation.json" in docker_web_job
    assert "artifacts/ci/pytest-docker-web.xml" in docker_web_job


def test_ci_pytest_runs_use_safe_cache_and_basetemp_boundaries() -> None:
    compose = (ROOT / "docker-compose.yml").read_text(encoding="utf-8")
    ci_check = (ROOT / "tools" / "ci_check.py").read_text(encoding="utf-8")

    assert "PYTEST_ADDOPTS: --basetemp=/tmp/pytest-tmp" in compose
    assert "artifacts/pytest-tmp" not in compose
    assert 'Path("/tmp/pytest-tmp").mkdir(parents=True, exist_ok=True)' in ci_check
    assert ci_check.count('"no:cacheprovider"') == 3
    assert 'docker_web_temp = ARTIFACTS / "docker-web-tmp"' in ci_check
    assert 'env["PYTEST_ADDOPTS"] = f"--basetemp={docker_web_temp}"' in ci_check
    assert ci_check.count("shutil.rmtree(docker_web_temp, ignore_errors=True)") == 2


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
