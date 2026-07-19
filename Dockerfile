ARG PYTHON_BASE=python:3.11-slim@sha256:e031123e3d85762b141ad1cbc56452ba69c6e722ebf2f042cc0dc86c47c0d8b3
FROM ${PYTHON_BASE}

ARG SETUPTOOLS_VERSION=83.0.0

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update \
    && apt-get install -y --no-install-recommends docker-cli \
    && rm -rf /var/lib/apt/lists/*

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XLSLIBERATOR_APPLICATION_CONTAINER=1 \
    XLSLIBERATOR_DATA_DIR=/data

WORKDIR /app
COPY pyproject.toml README.md ./
COPY src ./src
RUN python -m pip install --no-cache-dir --upgrade "setuptools==${SETUPTOOLS_VERSION}" \
    && python -m pip install --no-cache-dir ".[web]" \
    && cmp -s /app/src/sitecustomize.py \
        "$(python -c 'import site; print(site.getsitepackages()[0])')/sitecustomize.py"

RUN groupadd --gid 10001 appuser \
    && useradd \
        --uid 10001 \
        --gid 10001 \
        --create-home \
        --shell /usr/sbin/nologin \
        appuser \
    && mkdir -p /data \
    && chown -R appuser:appuser /data /app
USER 10001:10001

EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8080/healthz', timeout=3).read()"

CMD ["uvicorn", "xlsliberator.web.app:create_app", "--factory", "--host", "0.0.0.0", "--port", "8080"]
