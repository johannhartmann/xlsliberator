import os
import shutil
import subprocess

import pytest


@pytest.mark.integration
@pytest.mark.docker
def test_dockerfile_builds_when_enabled() -> None:
    if os.getenv("DOCKER_TESTS") != "1":
        pytest.skip("Set DOCKER_TESTS=1 to run Docker smoke tests")
    if shutil.which("docker") is None:
        pytest.skip("Docker is not installed")

    subprocess.run(
        ["docker", "build", "-t", "xlsliberator-web:test", "."],
        check=True,
        timeout=600,
    )
