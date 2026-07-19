# XLSLiberator Web App

The web app is an authenticated client of the repository's internal Open-SWE
service, the only supported agent and migration orchestrator. It accepts a
workbook, creates a private owner-scoped Open-SWE thread, streams safe stage
events, and downloads the independently reviewed delivery bundle.

The web container never imports the conversion API, starts LibreOffice,
connects to UNO, receives a Docker socket, or falls back to a local conversion
path. If Open-SWE is not configured or cannot be reached, migrations fail
closed.

The two bundled sample workbooks exercise upload, thread creation, progress, and
delivery. They are convenience inputs, not migration-quality evidence.

## Docker development

Docker is the only supported development and runtime platform. The host may run
Docker and file or Git operations only; never start host Python, PyUNO, UNO,
LibreOffice, or `soffice`.

Configure the embedded Open-SWE service:

```bash
cp .env.example .env
# Select one supported model and set only its matching provider key.
# Example: XLSLIBERATOR_OPEN_SWE_MODEL=openai:gpt-5.5
docker compose up -d --build xlsliberator-web
```

Open `http://127.0.0.1:8080/`. The published port is loopback-only by default.
The web service mounts only its `/data` volume.

This repository does not require or maintain a separate Open-SWE fork. The
`xlsliberator-open-swe` image verifies and installs a pinned upstream commit,
loads `xlsliberator.open_swe_agent.graph`, and serves the versioned workbook
routes under `/api/xlsliberator/migrations`. Open-SWE remains the sole owner of
agent state, model selection, specialist routing, and review. It can reach only
thread-confined files and the curated internal MCP tools; it receives neither a
shell backend nor the Docker socket.

`/readyz` reports:

- `data_dir_writable`
- `open_swe_configured`
- `open_swe_reachable`
- `target_libreoffice_version`

Traffic should be admitted only when the data directory is writable and
Open-SWE is reachable. Readiness never probes or launches a local office
executable.

## Configuration

- `XLSLIBERATOR_DATA_DIR`: private job root, default `/data`
- `XLSLIBERATOR_MAX_UPLOAD_MB`: upload limit, default `64`
- `XLSLIBERATOR_WEB_WORKERS`: bounded Open-SWE client concurrency, default `1`
- `XLSLIBERATOR_JOB_RETENTION_HOURS`: local delivery retention, default `24`
- `XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN`: shared web-to-agent bearer secret
- `XLSLIBERATOR_OPEN_SWE_MODEL`: required explicit model ID; blank starts no run
- `XLSLIBERATOR_OPEN_SWE_REASONING_EFFORT`: provider-supported effort
- `XLSLIBERATOR_OPEN_SWE_MAX_OUTPUT_TOKENS`: bounded model output
- `XLSLIBERATOR_OPEN_SWE_OWNER_ID`: stable tenant/user identity
- `XLSLIBERATOR_OPEN_SWE_POLL_SECONDS`: event/status polling interval
- `XLSLIBERATOR_OPEN_SWE_REQUEST_TIMEOUT_SECONDS`: individual API timeout
- `XLSLIBERATOR_OPEN_SWE_JOB_TIMEOUT_SECONDS`: whole migration time limit

The web client asks Open-SWE to delete private source and dependency copies
after completion. It also deletes its local private inputs before marking the
job complete. Downloaded deliverables expire through the local cleanup policy.

## API and no-JavaScript flow

- `GET /`: landing page and upload form
- `POST /jobs`: upload and redirect to the standalone job page
- `POST /api/jobs`: upload and return job JSON
- `GET /api/jobs/{job_id}`: safe status, Open-SWE thread ID, events, and artifacts
- `GET /api/jobs/{job_id}/events?since=0`: polling event stream
- `POST /api/jobs/{job_id}/messages`: follow-up on the same Open-SWE thread
- `POST /api/jobs/{job_id}/dependencies`: dependency upload on the same thread
- `POST /api/jobs/{job_id}/cancel`: Open-SWE cancellation
- `GET /api/jobs/{job_id}/artifacts`: delivery manifest
- `GET /jobs/{job_id}/artifacts/{artifact_id}`: owner-scoped deliverable
- `GET /jobs/{job_id}/download`: final ODS
- `GET /jobs/{job_id}/report.json`: migration report JSON
- `GET /jobs/{job_id}/report.md`: migration report Markdown
- `GET /healthz`, `GET /readyz`: liveness and dependency readiness

The standalone job page preserves form-based follow-up, dependency upload, and
cancellation for browsers without JavaScript.

## Publication boundary

Uploads and dependencies are untrusted. The web app validates filenames,
extensions, signatures, size, UUID job IDs, and safe download names. Public job
JSON contains no local paths. Open-SWE exposes only its owner-checked artifact
allowlist and must reject secrets, hidden-test material, private reasoning,
unsafe filenames, oversized artifacts, and internal paths.

Run the web app behind authenticated TLS with rate limiting and scanning. Keep
the bearer token in a secret manager and assign distinct owner IDs where tenant
isolation is required.

## Cleanup

```bash
docker compose run --rm xlsliberator-web \
  xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

## Docker smoke

```bash
mkdir -p artifacts/pytest-tmp artifacts/ci
docker compose --profile ci-runner run --rm test-runner \
  python tools/ci_check.py docker-web
```

The smoke starts the real in-repository MCP, Open-SWE, and web images under an
isolated Compose project. It assembles the actual Open-SWE graph and verifies
that the web reaches it. The model remains unset, so the test also proves that
upload execution fails closed without selecting a paid provider. No fake
Open-SWE HTTP server is involved.

A paid live-model migration remains explicitly opt-in. Start the embedded stack
with the chosen model/key in `.env`, then attach the test image to the same
Compose network:

```bash
docker compose up -d --build xlsliberator-web
docker run --rm --network xlsliberator_default --env-file .env \
  -e XLSLIBERATOR_ALLOW_PAID_OPEN_SWE_TEST=1 \
  -v "$PWD:/workspace" -w /workspace xlsliberator-test:py311 \
  pytest -m "integration and live" tests/integration/test_open_swe_real.py
```
