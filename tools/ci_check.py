#!/usr/bin/env python3
"""Run the same fail-closed checks used by GitHub Actions."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import cast

from xlsliberator.release_gates import GateAttestation, GateCheck, release_workspace_sha256

ROOT = Path(__file__).resolve().parents[1]
ARTIFACTS = ROOT / "artifacts" / "ci"
OFFICE_IMAGE = "xlsliberator-libreoffice:26.2.4.2"


def run(command: list[str], *, env: dict[str, str] | None = None) -> None:
    """Run one required command and stop on the first failure."""
    print("+", " ".join(command), flush=True)
    subprocess.run(command, cwd=ROOT, env=env, check=True)  # noqa: S603


def quality() -> None:
    """Run formatting, lint, typing, and office-free unit tests in this container."""
    run(["ruff", "format", "--check", "."])
    run(["ruff", "check", "."])
    run(["mypy", "src/"])
    run(
        [
            "pytest",
            "-p",
            "no:cacheprovider",
            "-m",
            "not integration",
            "--junitxml",
            str(ARTIFACTS / "pytest-unit.xml"),
        ]
    )


def office() -> None:
    """Build and test the pinned Docker-only LibreOffice runtime."""
    # The configured workspace policy resolves every root strictly.  Pytest
    # creates its basetemp lazily, so tests that do not request tmp_path still
    # need the Docker-only root to exist before backend discovery starts.
    Path("/tmp/pytest-tmp").mkdir(parents=True, exist_ok=True)
    run(
        [
            "docker",
            "build",
            "--file",
            "docker/office/libreoffice/Dockerfile",
            "--tag",
            OFFICE_IMAGE,
            ".",
        ]
    )
    run(["docker", "run", "--rm", "--network", "none", OFFICE_IMAGE, "runtime-probe"])
    env = dict(os.environ)
    env["XLSLIBERATOR_FAIL_ON_SKIP"] = "1"
    env["XLSLIBERATOR_RUNTIME_ARTIFACT_DIR"] = str(ARTIFACTS / "office-runtime")
    run(
        [
            "pytest",
            "-p",
            "no:cacheprovider",
            "tests/it/test_formula_parser_backend.py",
            "tests/it/test_real_libreoffice_conversion.py",
            "-m",
            "integration and docker and not live",
            "--junitxml",
            str(ARTIFACTS / "pytest-office.xml"),
        ],
        env=env,
    )


def docker_web() -> None:
    """Build and smoke-test the web image."""
    docker_web_temp = ARTIFACTS / "docker-web-tmp"
    shutil.rmtree(docker_web_temp, ignore_errors=True)
    docker_web_temp.mkdir(parents=True)
    env = dict(os.environ)
    env["DOCKER_TESTS"] = "1"
    env["XLSLIBERATOR_FAIL_ON_SKIP"] = "1"
    # Nested Docker bind mounts must originate from the host-visible checkout,
    # not from the test-orchestrator container's private /tmp filesystem.
    env["PYTEST_ADDOPTS"] = f"--basetemp={docker_web_temp}"
    try:
        run(
            [
                "pytest",
                "-p",
                "no:cacheprovider",
                "tests/integration/test_docker_web.py",
                "--junitxml",
                str(ARTIFACTS / "pytest-docker-web.xml"),
            ],
            env=env,
        )
    finally:
        shutil.rmtree(docker_web_temp, ignore_errors=True)


def package() -> None:
    """Build and validate distribution artifacts."""
    # The test image contains the declared build backend.  Avoid an implicit
    # PyPI fetch so the package gate remains reproducible with network disabled.
    run([sys.executable, "-m", "build", "--no-isolation"])
    distributions = sorted(
        str(path) for pattern in ("*.whl", "*.tar.gz") for path in (ROOT / "dist").glob(pattern)
    )
    if not distributions:
        raise RuntimeError("Package build produced no distributions")
    wheels = [Path(path) for path in distributions if path.endswith(".whl")]
    if not wheels:
        raise RuntimeError("Package build produced no wheel")
    expected_guard = (ROOT / "src" / "sitecustomize.py").read_bytes()
    for wheel in wheels:
        with zipfile.ZipFile(wheel) as archive:
            try:
                packaged_guard = archive.read("sitecustomize.py")
            except KeyError as exc:
                raise RuntimeError(f"{wheel.name} omits the host-UNO startup guard") from exc
        if packaged_guard != expected_guard:
            raise RuntimeError(f"{wheel.name} contains an invalid host-UNO startup guard")
    run([sys.executable, "-m", "twine", "check", *distributions])


def security() -> None:
    """Run required dependency and source security checks."""
    run(["pip-audit", "--desc"])
    run(["bandit", "-r", "src/", "-c", "pyproject.toml"])


CHECKS = {
    "quality": quality,
    "office": office,
    "docker-web": docker_web,
    "package": package,
    "security": security,
}


def write_attestation(check: str) -> None:
    """Record a gate only after every subprocess in that gate returned zero."""
    if (
        os.environ.get("XLSLIBERATOR_APPLICATION_CONTAINER") != "1"
        or not Path("/.dockerenv").is_file()
    ):
        raise RuntimeError("CI attestations may be produced only in the Docker test image")
    path = ARTIFACTS / f"{check}-attestation.json"
    attestation = GateAttestation(
        check=cast(GateCheck, check),
        status="passed",
        runner="xlsliberator-test-container",
        workspace_sha256=release_workspace_sha256(ROOT),
        recorded_at=datetime.now(UTC),
    )
    path.write_text(attestation.model_dump_json(indent=2) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("check", choices=[*CHECKS, "all"], nargs="?", default="all")
    args = parser.parse_args()
    ARTIFACTS.mkdir(parents=True, exist_ok=True)
    selected = CHECKS.items() if args.check == "all" else [(args.check, CHECKS[args.check])]
    for name, check in selected:
        check()
        write_attestation(name)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
