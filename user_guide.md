# XLSLiberator Web User Guide

## What the application does

The web app accepts Excel workbooks and delegates migrations to Open-SWE.
Open-SWE is the only supported agent and migration orchestrator. XLSLiberator
supplies deterministic workbook tools and the pinned LibreOffice target runtime;
the web container does not convert workbooks locally.

If Open-SWE is absent, unconfigured, or unreachable, the migration fails closed.
The Open-SWE runtime is embedded in this repository's Docker stack. There is no
alternate agent, second repository, or local conversion fallback.

## Start the app

Docker is the only supported platform. The host may run Docker, Git, and file
operations only.

```bash
cp .env.example .env
# Select one supported model and set only its matching provider key.
# Example: XLSLIBERATOR_OPEN_SWE_MODEL=openai:gpt-5.5
docker compose up -d --build xlsliberator-web
```

Open `http://127.0.0.1:8080/`. The port is published on loopback only.

## Basic workflow

1. Upload an `.xls`, `.xlsx`, `.xlsm`, or `.xlsb` workbook.
2. XLSLiberator creates an owner-scoped job and an Open-SWE migration thread.
3. Follow sanitized stage events on the job page.
4. Add requirements or dependency files to the same thread when needed.
5. Download only the reviewed artifacts published by Open-SWE.

The web app generates server-side job IDs, stores inputs outside the web root,
checks extensions and signatures, enforces size limits, and never exposes local
filesystem paths.

## Progress and follow-up

The job page polls:

```text
GET /api/jobs/{job_id}/events?since=0
```

Follow-up requirements and dependency files remain attached to the same
Open-SWE thread:

```text
POST /api/jobs/{job_id}/messages
POST /api/jobs/{job_id}/dependencies
```

Cancellation is propagated to Open-SWE:

```text
POST /api/jobs/{job_id}/cancel
```

## Downloads and reports

Completed jobs may expose:

```text
GET /jobs/{job_id}/download
GET /jobs/{job_id}/report.json
GET /jobs/{job_id}/report.md
GET /api/jobs/{job_id}/artifacts
GET /jobs/{job_id}/artifacts/{artifact_id}
```

Artifact IDs and filenames come from an owner-checked allowlist. A converted
file alone is not proof of formula, macro, UI, or behavioral equivalence; the
delivery must include the applicable Open-SWE review and XLSLiberator evidence.

## JSON API

Create a job:

```bash
curl -F "file=@workbook.xlsx" \
  -H "Accept: application/json" \
  http://127.0.0.1:8080/api/jobs
```

Check status or poll events:

```bash
curl http://127.0.0.1:8080/api/jobs/<job_id>
curl "http://127.0.0.1:8080/api/jobs/<job_id>/events?since=0"
```

## Configuration

| Variable | Default | Purpose |
| --- | --- | --- |
| `XLSLIBERATOR_DATA_DIR` | `/data` | Private job and artifact root |
| `XLSLIBERATOR_MAX_UPLOAD_MB` | `64` | Maximum accepted upload size |
| `XLSLIBERATOR_WEB_WORKERS` | `1` | Bounded Open-SWE client concurrency |
| `XLSLIBERATOR_JOB_RETENTION_HOURS` | `24` | Local delivery retention |
| `XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN` | local-only placeholder | Internal bearer secret |
| `XLSLIBERATOR_OPEN_SWE_MODEL` | unset | Explicit Open-SWE model; blank starts no run |
| `XLSLIBERATOR_OPEN_SWE_REASONING_EFFORT` | `medium` | Provider-supported effort |
| `XLSLIBERATOR_GITHUB_MODELS_ENABLED` | `0` | Additional explicit GitHub Models opt-in |
| `XLSLIBERATOR_OPEN_SWE_OWNER_ID` | `xlsliberator-web` | Stable owner identity |
| `XLSLIBERATOR_OPEN_SWE_POLL_SECONDS` | `1` | Status polling interval |
| `XLSLIBERATOR_OPEN_SWE_REQUEST_TIMEOUT_SECONDS` | `60` | Per-request timeout |
| `XLSLIBERATOR_OPEN_SWE_JOB_TIMEOUT_SECONDS` | `3600` | Whole-migration timeout |

Provider and model selection are explicit Open-SWE operator decisions. The web
container has no provider credential or SDK and cannot select GitHub Models or
make any provider a gate for deterministic XLSLiberator operations.

## Health checks

`GET /healthz` confirms that the FastAPI process responds. `GET /readyz`
separately reports whether local storage is writable and whether Open-SWE is
configured and reachable. Readiness never discovers or starts a local
LibreOffice executable.

## Cleanup and security

```bash
docker compose run --rm xlsliberator-web \
  xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

The web client asks Open-SWE to delete private source and dependency copies
after completion and deletes its own private inputs before publishing the job.
Run the service behind authenticated TLS with rate limiting, upload scanning,
resource limits, and operational monitoring before any remote exposure.

For deployment and test details, see [docs/web_app.md](docs/web_app.md).
