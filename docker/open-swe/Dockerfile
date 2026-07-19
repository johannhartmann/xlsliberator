ARG PYTHON_BASE=python:3.11-slim@sha256:e031123e3d85762b141ad1cbc56452ba69c6e722ebf2f042cc0dc86c47c0d8b3
FROM ${PYTHON_BASE}

ARG OPEN_SWE_COMMIT=f0897479c38f2506f03b4de38081d4770928f09d
ARG OPEN_SWE_ARCHIVE_SHA256=4f5d639f36c1f68d5ddb47a90ee83f83be5b233f1b50bfac92dff671674b3fb0
ARG SETUPTOOLS_VERSION=83.0.0
ARG UV_VERSION=0.9.26

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XLSLIBERATOR_APPLICATION_CONTAINER=1 \
    LANGGRAPH_CLI_NO_ANALYTICS=1 \
    LANGGRAPH_NO_VERSION_CHECK=1 \
    OPEN_SWE_UPSTREAM_COMMIT=${OPEN_SWE_COMMIT}

RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /opt
RUN curl -fsSL \
      "https://codeload.github.com/langchain-ai/open-swe/tar.gz/${OPEN_SWE_COMMIT}" \
      -o /tmp/open-swe.tar.gz \
    && echo "${OPEN_SWE_ARCHIVE_SHA256}  /tmp/open-swe.tar.gz" | sha256sum -c - \
    && mkdir /opt/open-swe \
    && tar -xzf /tmp/open-swe.tar.gz --strip-components=1 -C /opt/open-swe \
    && rm /tmp/open-swe.tar.gz

WORKDIR /app
COPY pyproject.toml README.md langgraph.open-swe.json ./
COPY src ./src
RUN python -m pip install --no-cache-dir --upgrade \
      "setuptools==${SETUPTOOLS_VERSION}" \
      "uv==${UV_VERSION}" \
    && cd /opt/open-swe \
    && uv sync --frozen --no-dev \
    && uv pip install --python /opt/open-swe/.venv/bin/python "/app[web]"

ENV PATH="/opt/open-swe/.venv/bin:${PATH}"

RUN groupadd --gid 10001 appuser \
    && useradd \
        --uid 10001 \
        --gid 10001 \
        --create-home \
        --shell /usr/sbin/nologin \
        appuser \
    && mkdir -p /workspaces /tmp/open-swe \
    && chown -R appuser:appuser /workspaces /tmp/open-swe /app

USER 10001:10001
EXPOSE 2024
HEALTHCHECK --interval=15s --timeout=5s --start-period=30s --retries=5 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:2024/health', timeout=3).read()"

CMD ["langgraph", "dev", "--no-browser", "--no-reload", "--host", "0.0.0.0", "--port", "2024", "--config", "/app/langgraph.open-swe.json"]
