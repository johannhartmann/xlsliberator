# XLSLiberator Web App

The web app serves a marketing landing page (`GET /`) with an embedded **live demo**:
visitors pick a bundled example workbook or upload their own, start a real conversion,
watch the pipeline progress inline, and download the converted `.ods` file plus JSON and
Markdown reports — without leaving the page. The standalone job page (`/jobs/{job_id}`)
remains available as a permalink and as the no-JavaScript fallback target.

## Docker Development

Host Python and host office execution are forbidden. Build and run the application
through Docker Compose only:

```bash
mkdir -p artifacts/runtime-tmp
docker compose build libreoffice-runtime xlsliberator-web
docker compose up -d xlsliberator-web
```

Open `http://localhost:8080/`. The published port is loopback-only by default.

## Docker Compose

The service listens on `http://localhost:8080` and stores job artifacts in the
`xlsliberator-data` named volume. The trusted web container has the Docker CLI
and Docker-socket access only so it can create a disposable office worker. The
worker receives only read-only input and isolated writable job mounts; it never
receives the Docker socket. `/readyz` must report the pinned Docker office runtime
as available before an operator sends traffic. Never verify readiness by starting
`soffice` in the web container or on the host.

```bash
docker compose exec xlsliberator-web python -c \
  "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8080/readyz').read()"
```

## Configuration

- `XLSLIBERATOR_DATA_DIR`: artifact root, default `/data`
- `XLSLIBERATOR_RUNTIME_TEMP_ROOT`: web-container staging root, default from Compose
- `XLSLIBERATOR_DOCKER_HOST_RUNTIME_TEMP_ROOT`: matching host path used for nested
  Docker bind mounts; Compose derives it from `${PWD}`
- `DOCKER_GID`: Docker-socket group ID on Linux, default `0`; set it to the output
  of `stat -c %g /var/run/docker.sock` when the daemon socket uses another group
- `XLSLIBERATOR_MAX_UPLOAD_MB`: max upload size, default `100`
- `XLSLIBERATOR_WEB_WORKERS`: conversion concurrency, default `1`
- `XLSLIBERATOR_JOB_RETENTION_HOURS`: cleanup age, default `24`
- `ANTHROPIC_API_KEY`: optional; without it, macro translation uses fallback behavior

## API

- `GET /`: marketing landing page with the embedded live demo
- `POST /jobs`: browser upload, redirects to `/jobs/{job_id}` (no-JS fallback)
- `POST /api/jobs`: JSON upload, returns job status (used by the inline demo)
- `GET /api/jobs/{job_id}`: status JSON
- `GET /api/jobs/{job_id}/events?since=0`: polling progress events
- `GET /api/jobs/{job_id}/report`: conversion report summary JSON (for the inline demo)
- `POST /api/jobs/{job_id}/cancel`: best-effort cancellation
- `GET /jobs/{job_id}/download`: converted `.ods`
- `GET /jobs/{job_id}/report.json`: conversion report JSON
- `GET /jobs/{job_id}/report.md`: conversion report Markdown
- `GET /healthz`, `GET /readyz`: health and readiness

## Security Notes

Uploads are treated as hostile. The app validates extensions and basic file signatures,
generates UUID job IDs, stores files outside the webroot, and never exposes internal
paths in JSON responses. The web process is a privileged orchestration boundary because
it can reach the Docker daemon; it must not be treated as an untrusted workbook runtime.
The web path disables global LibreOffice macro security changes and creates per-job
LibreOffice profile directories. Do not expose the service publicly without a separate
authenticated gateway, scanning, rate limits, TLS, and preferably a remote or microVM
worker boundary.

Legacy `.xls` parsing remains incomplete and should not be represented as fully validated. Macro-heavy workbooks may need manual review, especially when no `ANTHROPIC_API_KEY` is configured.

## Cleanup

Old job directories are removed on startup according to `XLSLIBERATOR_JOB_RETENTION_HOURS`. Manual cleanup:

```bash
docker compose run --rm xlsliberator-web \
  xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

## Blocking Docker Smoke Test

```bash
mkdir -p artifacts/runtime-tmp artifacts/pytest-tmp artifacts/ci
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  python tools/ci_check.py docker-web
```

The test builds both images, starts the web container, requires `/readyz` to prove
the pinned `26.2.4.2` runtime, uploads a generated XLSX, waits for a completed job,
and verifies the produced ODS package. CI sets `XLSLIBERATOR_FAIL_ON_SKIP=1`, so an
unexpected skip or missing Docker runtime fails the job.
