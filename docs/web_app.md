# XLSLiberator Web App

The web app serves a marketing landing page (`GET /`) with an embedded **live demo**:
visitors pick a bundled example workbook or upload their own, start a real conversion,
watch the pipeline progress inline, and download the converted `.ods` file plus JSON and
Markdown reports — without leaving the page. The standalone job page (`/jobs/{job_id}`)
remains available as a permalink and as the no-JavaScript fallback target.

## Local Development

Install the optional web dependencies:

```bash
pip install -e ".[web,dev]"
```

Run the app:

```bash
uvicorn xlsliberator.web.app:create_app --factory --reload
# or
xlsliberator web-serve --host 0.0.0.0 --port 8080 --reload
```

Open `http://localhost:8080/`.

## Docker Compose

```bash
docker compose up --build
```

The service listens on `http://localhost:8080` and stores job artifacts in the `xlsliberator-data` named volume. To verify the image manually:

```bash
docker build -t xlsliberator-web .
docker run --rm xlsliberator-web soffice --version
```

## Configuration

- `XLSLIBERATOR_DATA_DIR`: artifact root, default `/data`
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

Uploads are treated as hostile. The app validates extensions and basic file signatures, generates UUID job IDs, stores files outside the webroot, and never exposes internal paths in JSON responses. The web path disables global LibreOffice macro security changes and creates per-job LibreOffice profile directories. Do not expose the service publicly without authentication, scanning, rate limits, and TLS.

Legacy `.xls` parsing remains incomplete and should not be represented as fully validated. Macro-heavy workbooks may need manual review, especially when no `ANTHROPIC_API_KEY` is configured.

## Cleanup

Old job directories are removed on startup according to `XLSLIBERATOR_JOB_RETENTION_HOURS`. Manual cleanup:

```bash
xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

## Optional Docker Smoke Test

```bash
DOCKER_TESTS=1 uv run pytest tests/integration/test_docker_web.py
```

The test skips when Docker is unavailable or `DOCKER_TESTS` is not set.
In GitHub Actions, run the `CI` workflow manually with `docker-smoke=true` to build the image through the same smoke test.
