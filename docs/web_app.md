# XLSLiberator Web App

The web app is an authenticated client of the Open-SWE workbook-migration API.
It accepts a workbook, creates a private owner-scoped thread, streams safe stage
events, and downloads the independently reviewed delivery bundle. It never
imports the conversion API, starts LibreOffice, connects to UNO, or receives a
Docker socket.

The two bundled sample workbooks are **Level 0 basic-pipeline examples**. They
exercise upload, thread creation, progress, and delivery, but are not evidence
for serious macro, UI, external-dependency, XLS/XLSB, or hostile-workbook
migrations. Representative migration episodes live in the corpus described by
the agentic implementation roadmap.

## Docker Development

Docker is the only supported development and runtime platform. The host may run
Docker and file/Git operations only; never start host Python, PyUNO, UNO,
LibreOffice, or `soffice`.

Configure the Open-SWE endpoint and its shared trigger credential:

```bash
cp .env.example .env
# Set XLSLIBERATOR_OPEN_SWE_URL and XLSLIBERATOR_OPEN_SWE_TOKEN in .env.
docker compose build xlsliberator-web
docker compose up -d xlsliberator-web
```

Open `http://localhost:8080/`. The published port is loopback-only by default.
The web service mounts only its `/data` volume. LibreOffice `26.2.4.2`, its
bundled Python, UNO, and PyUNO remain inside the Open-SWE-managed pinned office
runtime.

`/readyz` reports ready inputs separately:

- `data_dir_writable`
- `open_swe_configured`
- `open_swe_reachable`
- `target_libreoffice_version`

An operator should send traffic only when both writable and reachable are true.
Readiness never probes or launches a local office executable.

## Configuration

- `XLSLIBERATOR_DATA_DIR`: private job root, default `/data`
- `XLSLIBERATOR_MAX_UPLOAD_MB`: upload limit, default `64`
- `XLSLIBERATOR_WEB_WORKERS`: bounded thread-client concurrency, default `1`
- `XLSLIBERATOR_JOB_RETENTION_HOURS`: local delivery retention, default `24`
- `XLSLIBERATOR_OPEN_SWE_URL`: Open-SWE API base URL
- `XLSLIBERATOR_OPEN_SWE_TOKEN`: bearer credential, supplied as a secret
- `XLSLIBERATOR_OPEN_SWE_OWNER_ID`: stable tenant/user identity
- `XLSLIBERATOR_OPEN_SWE_POLL_SECONDS`: event/status polling interval
- `XLSLIBERATOR_OPEN_SWE_REQUEST_TIMEOUT_SECONDS`: individual API timeout
- `XLSLIBERATOR_OPEN_SWE_JOB_TIMEOUT_SECONDS`: whole migration time limit

The web client asks Open-SWE to delete private source/dependency copies after
completion. It also deletes its own source/dependency copies before marking the
job complete. Downloaded deliverables expire through the local cleanup policy.

## API and no-JavaScript flow

- `GET /`: landing page with Level 0 sample workbooks and upload form
- `POST /jobs`: upload and redirect to the standalone job page
- `POST /api/jobs`: upload and return job JSON
- `GET /api/jobs/{job_id}`: safe status, thread ID, events, and artifact metadata
- `GET /api/jobs/{job_id}/events?since=0`: polling event stream
- `POST /api/jobs/{job_id}/messages`: follow-up requirement on the same thread
- `POST /api/jobs/{job_id}/dependencies`: dependency upload on the same thread
- `POST /api/jobs/{job_id}/cancel`: remote cancellation
- `GET /api/jobs/{job_id}/artifacts`: delivery manifest
- `GET /jobs/{job_id}/artifacts/{artifact_id}`: one owner-scoped deliverable
- `GET /jobs/{job_id}/download`: final ODS
- `GET /jobs/{job_id}/report.json`: migration report JSON
- `GET /jobs/{job_id}/report.md`: migration report Markdown
- `GET /healthz`, `GET /readyz`: liveness and dependency readiness

The standalone job page preserves form-based follow-up, dependency upload, and
cancellation for browsers without JavaScript.

## Publication boundary

Uploads and dependencies are untrusted. The web app validates filenames,
extensions, signatures, size, UUID job IDs, and safe download names. Public job
JSON contains no local paths. Open-SWE exposes only a fixed deliverable
allowlist and rejects credentials, hidden-test material, system prompts, private
reasoning, unsafe filenames, oversized artifacts, and binary artifacts carrying
internal paths. Text artifacts have internal paths redacted.

Run the web app behind authenticated TLS with rate limiting and scanning. Keep
the bearer token in a secret manager and assign distinct owner IDs where tenant
isolation is required.

## Cleanup

Old job directories are removed on startup:

```bash
docker compose run --rm xlsliberator-web \
  xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

## Blocking Docker smoke

```bash
mkdir -p artifacts/pytest-tmp artifacts/ci
docker compose --profile ci-orchestrator run --rm test-orchestrator \
  python tools/ci_check.py docker-web
```

The smoke starts an in-container fake Open-SWE HTTP service and a separate web
container without Docker-socket access. It verifies authentication headers,
thread creation, safe stage streaming, ODS and evidence delivery, readiness,
and private-source deletion. A real LangGraph migration remains an explicitly
environment-gated integration because it requires the deployed Open-SWE graph,
sandbox, and pinned office services:

```bash
docker compose run --rm \
  -e XLSLIBERATOR_REAL_OPEN_SWE_URL \
  -e XLSLIBERATOR_REAL_OPEN_SWE_TOKEN \
  -e XLSLIBERATOR_REAL_OPEN_SWE_OWNER_ID \
  test pytest -m "integration and live" tests/integration/test_open_swe_real.py
```
