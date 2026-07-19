ARG PYTHON_BASE=python:3.11-slim@sha256:e031123e3d85762b141ad1cbc56452ba69c6e722ebf2f042cc0dc86c47c0d8b3
FROM ${PYTHON_BASE}

ARG PIP_VERSION=26.1.2
ARG SETUPTOOLS_VERSION=83.0.0

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PYTHONPATH=/workspace/src \
    XLSLIBERATOR_APPLICATION_CONTAINER=1 \
    XLSLIBERATOR_OFFICE_CONTAINER=0 \
    HOME=/tmp/home

RUN apt-get update \
    && apt-get install -y --no-install-recommends docker-cli docker-compose git zip \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /build
COPY pyproject.toml README.md ./
COPY src ./src
RUN python -m pip install --no-cache-dir --upgrade \
        "pip==${PIP_VERSION}" \
        "setuptools==${SETUPTOOLS_VERSION}" \
    && python -m pip install --no-cache-dir ".[web,dev]" \
    && python -m pip install --no-cache-dir \
        bandit[toml] build hatchling pip-audit pre-commit pytest-cov twine \
    && cmp -s /build/src/sitecustomize.py \
        "$(python -c 'import site; print(site.getsitepackages()[0])')/sitecustomize.py"

RUN useradd --uid 10001 --create-home --shell /usr/sbin/nologin tester \
    && mkdir -p /workspace /tmp/home \
    && chown -R tester:tester /workspace /tmp/home /build

USER 10001:10001
WORKDIR /workspace
CMD ["pytest", "-m", "not integration"]
