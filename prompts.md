
# Concept: Docker-based XLSLiberator Web App

## Product behavior

The user experience should be:

1. User opens `/`.
2. User uploads `.xls`, `.xlsx`, `.xlsm`, or `.xlsb`.
3. The app creates a job and immediately shows a progress page.
4. The progress page streams or polls structured status events:

   * uploaded
   * queued
   * analyzing workbook
   * converting with LibreOffice
   * repairing formulas and named ranges
   * extracting VBA
   * translating macros
   * embedding macros
   * verifying formulas
   * verifying macros
   * verifying GUI/events, where available
   * complete or failed
5. The user can download:

   * resulting `.ods` OpenOffice/LibreOffice Calc file
   * conversion report JSON
   * human-readable report Markdown or HTML
   * optional log bundle for debugging

The file delivered should be `.ods`, because that is the OpenDocument spreadsheet format produced by the current converter and supported by OpenOffice Calc.

## Recommended stack

Use **FastAPI** for the web/API layer. It fits the existing Python/Pydantic style and supports file uploads through `UploadFile`; FastAPI’s docs note that file uploads require `python-multipart`, and `UploadFile` uses a spooled file that moves larger data to disk instead of reading everything into memory. ([FastAPI][1])

Use **Jinja2 + a small amount of browser JavaScript** for the first version. A full React/Vue frontend would slow down implementation and create a second build pipeline. FastAPI can return HTML responses directly, and it provides `FileResponse` for async file downloads. ([FastAPI][2])

Use **Server-Sent Events or polling** for progress. SSE is ideal for a one-way “job status log” stream; FastAPI supports streaming responses with async generators through `StreamingResponse`. ([FastAPI][2])

Use an **in-process job manager first**, with a strict single-worker or low-concurrency executor. LibreOffice/UNO conversion is heavy and not naturally safe to run many conversions in parallel in the same profile. A future production version can move to Redis/RQ/Celery, but the first Docker app can be reliable with `ThreadPoolExecutor(max_workers=1)` or a worker subprocess.

Use Docker Compose for local operation. FastAPI’s official Docker guide recommends copying dependency definitions before app code to improve build cache behavior and using exec-form `CMD` so shutdown signals work correctly. ([FastAPI][3]) Compose named volumes are suitable for persistent job artifacts because Docker Compose volumes are persistent stores that must be explicitly granted to services. ([Docker Documentation][4])

## Security model

This app accepts hostile files by design, so treat every upload as unsafe. OWASP’s upload guidance recommends allowlisted extensions, not trusting user-provided `Content-Type`, generating server-side filenames, enforcing file size limits, storing files outside the webroot, and applying defense-in-depth validation. ([OWASP Cheat Sheet Series][5])

For this repo, that means:

* Accept only `.xls`, `.xlsx`, `.xlsm`, `.xlsb`.
* Generate a UUID job ID and never use the uploaded filename as a path.
* Store uploads/results under `/data/jobs/<job_id>/`.
* Enforce max upload size, for example 50–100 MB initially.
* Reject path traversal and double-extension tricks.
* Never expose raw filesystem paths.
* Download only through `/jobs/{job_id}/download` after checking job state.
* Run LibreOffice with isolated temp/profile directories per job.
* Disable global macro security mutation from the web flow.
* Use a non-root container user if LibreOffice works with it.
* Add cleanup/retention, for example delete jobs older than 24 hours.

## Proposed repository layout

```text
src/xlsliberator/
  web/
    __init__.py
    app.py                 # FastAPI app factory
    routes.py              # HTML + JSON routes
    jobs.py                # Job model, store, executor
    progress.py            # Progress event bus
    schemas.py             # Pydantic API models
    security.py            # file validation, safe names, limits
    templates/
      base.html
      index.html
      job.html
      error.html
    static/
      app.css
      app.js

docker/
  entrypoint.sh

Dockerfile
docker-compose.yml
.dockerignore
docs/web_app.md
tests/unit/web/
tests/integration/web/
```

## Web API shape

```text
GET  /
  HTML upload form

POST /jobs
  multipart upload
  returns redirect to /jobs/{job_id} or JSON {job_id, status_url}

GET /jobs/{job_id}
  HTML progress page

GET /api/jobs/{job_id}
  JSON job status

GET /api/jobs/{job_id}/events
  text/event-stream status/progress stream

GET /jobs/{job_id}/download
  downloads converted .ods

GET /jobs/{job_id}/report.json
  downloads ConversionReport JSON

GET /jobs/{job_id}/report.md
  downloads ConversionReport Markdown

GET /healthz
  basic health

GET /readyz
  checks that soffice is discoverable and work directory is writable
```

## Job state model

```python
class JobPhase(str, Enum):
    UPLOADED = "uploaded"
    QUEUED = "queued"
    ANALYZING = "analyzing"
    CONVERTING = "converting"
    TRANSLATING = "translating"
    VERIFYING = "verifying"
    COMPLETED = "completed"
    FAILED = "failed"
    CANCELLED = "cancelled"

class JobEvent(BaseModel):
    job_id: str
    phase: JobPhase
    step: str
    message: str
    percent: int | None = None
    level: Literal["info", "warning", "error"] = "info"
    timestamp: datetime
    details: dict[str, Any] = Field(default_factory=dict)

class WebJob(BaseModel):
    id: str
    original_filename: str
    safe_input_path: Path
    output_path: Path
    report_json_path: Path
    report_md_path: Path
    status: JobPhase
    events: list[JobEvent]
    created_at: datetime
    updated_at: datetime
    error: str | None = None
```

## Progress instrumentation

The current `convert()` function logs many step names but does not expose a progress callback. The clean implementation is to add a new optional callback without breaking public API:

```python
ProgressCallback = Callable[[str, str, dict[str, Any]], None]

def convert(
    input_path: str | Path,
    output_path: str | Path,
    *,
    locale: str = "en-US",
    strict: bool = False,
    embed_macros: bool = True,
    use_agent: bool = True,
    progress_callback: ProgressCallback | None = None,
) -> ConversionReport:
    ...
```

Then the web worker emits events around existing stages:

```python
progress("analyzing", "Extracting workbook metadata", {})
progress("converting", "Running LibreOffice native conversion", {})
progress("translating", "Extracting and translating VBA modules", {})
progress("verifying", "Comparing formulas and validating macros", {})
progress("completed", "Conversion complete", {"output": "...ods"})
```

If you want zero changes to `api.py` in the first PR, the job worker can emit coarse progress before and after `convert()` and parse the final report afterward. But the best UX requires an explicit progress callback.

## Docker image concept

Base image:

```Dockerfile
FROM python:3.11-slim
```

Install LibreOffice packages:

```text
libreoffice
libreoffice-calc
libreoffice-script-provider-python
python3-uno
fonts-dejavu
```

Install the app:

```text
pip install -e ".[web]"
```

Run:

```text
uvicorn xlsliberator.web.app:create_app --factory --host 0.0.0.0 --port 8080
```

Add a healthcheck that hits `/healthz` and `/readyz`.

## Compose concept

```yaml
services:
  xlsliberator-web:
    build: .
    ports:
      - "8080:8080"
    environment:
      XLSLIBERATOR_DATA_DIR: /data
      XLSLIBERATOR_MAX_UPLOAD_MB: "100"
      ANTHROPIC_API_KEY: "${ANTHROPIC_API_KEY:-}"
    volumes:
      - xlsliberator-data:/data
    restart: unless-stopped

volumes:
  xlsliberator-data:
```

## Minimal viable implementation plan

### Milestone 1: Docker can run the existing CLI

Add `Dockerfile`, `.dockerignore`, `docker-compose.yml`, and docs. The container should run `soffice --version`, `python -m xlsliberator.cli --help`, and a tiny non-macro conversion fixture.

### Milestone 2: Web upload + job store

Add FastAPI app, templates, file upload validation, UUID job directories, and an in-memory job store. Uploading a valid Excel file creates a job and redirects to a progress page.

### Milestone 3: Background conversion worker

Add `WebJobRunner` that calls `convert()`, writes `.ods`, `report.json`, and `report.md`, and updates job status.

### Milestone 4: Real progress events

Add `progress_callback` to `convert()` and emit stage events from the core pipeline.

### Milestone 5: Verification UX

Show formulas matched/mismatched, macro modules validated, macro execution pass/fail, warnings, and errors using `ConversionReport`.

### Milestone 6: Hardening

Add upload limits, cleanup, isolated LibreOffice profiles, non-root container, tests, CI Docker build, and docs.

---

# Codex Prompt Pack

Use these as separate Codex tasks. Do not run overlapping tasks that edit the same files at the same time.

## Global Codex context

```text
You are working in the johannhartmann/xlsliberator repository.

Goal:
Add a Docker-based browser web app where a user can upload an Excel file, watch analysis/translation/verification progress, and download the resulting OpenOffice/LibreOffice Calc .ods file plus reports.

Current repo facts:
- The core API is src/xlsliberator/api.py: convert(input_path, output_path, locale="en-US", strict=False, embed_macros=True, use_agent=True) -> ConversionReport.
- The CLI is src/xlsliberator/cli.py with convert and mcp-serve commands.
- pyproject.toml has no FastAPI/Uvicorn/Jinja2/python-multipart dependencies yet.
- The repo already uses Pydantic v2, Click, Loguru, FastMCP, and pytest.
- CI runs ruff, mypy, unit tests, and LibreOffice integration tests.
- Existing conversion code currently changes global LibreOffice macro security. Do not rely on that for the web app.
- .xls parsing is incomplete and must not be represented as fully validated.
- Keep existing public APIs backward-compatible unless a prompt explicitly extends them.

Implementation constraints:
- Add tests for each feature.
- Unit tests must not require LibreOffice, Docker, or an Anthropic key.
- Integration tests that require LibreOffice or Docker must be marked and skippable.
- Avoid storing user files in the webroot.
- Validate upload extensions and size.
- Generate server-side job IDs and filenames.
- Never trust uploaded filenames or Content-Type.
- Do not expose raw filesystem paths in API responses.
- Prefer small, reviewable commits.

Before coding:
1. Inspect relevant files.
2. Produce a short plan.
3. Implement the smallest useful slice.
4. Run focused tests and lint where possible.
5. Report changed files and test results.
```

---

## Prompt 1 — Add web dependency extras

```text
Task:
Add a `web` optional dependency group to pyproject.toml.

Add dependencies:
- fastapi[standard] with a sensible lower bound
- uvicorn[standard] if not already included by fastapi standard in a way that works for this project
- jinja2
- python-multipart
- aiofiles, only if needed by implementation
- anyio, only if directly used

Also update mypy missing import overrides if needed.

Acceptance criteria:
- Existing default install remains unchanged except optional extras.
- `pip install -e ".[web,dev]"` should be the documented install for web development.
- No web code yet unless necessary.
- Update README or docs/web_app.md stub with install command.
- Run pyproject validation/build metadata check if available.
```

---

## Prompt 2 — Add Dockerfile and compose skeleton

```text
Task:
Add Docker support for running XLSLiberator in a container.

Files to add:
- Dockerfile
- docker-compose.yml
- .dockerignore
- docs/web_app.md
- docker/entrypoint.sh if useful

Docker requirements:
- Base on python:3.11-slim.
- Install LibreOffice Calc and Python UNO bridge packages on Debian/Ubuntu.
- Install the package with web extras: pip install -e ".[web]".
- Create /data for job artifacts.
- Expose port 8080.
- Use exec-form CMD.
- Add a simple healthcheck that can work once the web app exists. If web app does not exist yet, document it as TODO or use a harmless command like python -c import xlsliberator.
- Compose should mount a named volume at /data.
- Compose should pass ANTHROPIC_API_KEY optionally.

Acceptance criteria:
- `docker build .` succeeds.
- `docker run --rm IMAGE soffice --version` or equivalent documented command works.
- Docs explain how to run `docker compose up`.
- No web routes yet.
```

---

## Prompt 3 — Create FastAPI app skeleton

```text
Task:
Add a minimal FastAPI web app skeleton.

Files:
- src/xlsliberator/web/__init__.py
- src/xlsliberator/web/app.py
- src/xlsliberator/web/routes.py
- src/xlsliberator/web/templates/base.html
- src/xlsliberator/web/templates/index.html
- src/xlsliberator/web/static/app.css
- src/xlsliberator/web/static/app.js
- tests/unit/web/test_app.py

Behavior:
- create_app() returns a FastAPI instance.
- GET / returns an HTML upload page.
- GET /healthz returns {"status": "ok"}.
- GET /readyz returns JSON with:
  - data_dir_writable
  - soffice_available
  - version if discoverable
- Mount static files.
- Use Jinja2 templates.

Do not implement uploads yet.

Acceptance criteria:
- FastAPI TestClient tests pass without LibreOffice installed.
- /readyz must not fail if soffice is missing; it should report unavailable.
- Add docs command: uvicorn xlsliberator.web.app:create_app --factory --reload.
```

---

## Prompt 4 — Add CLI command `web-serve`

```text
Task:
Add a CLI command to serve the web app.

Modify:
- src/xlsliberator/cli.py

Add:
- `xlsliberator web-serve --host 0.0.0.0 --port 8080 --reload`

Behavior:
- Imports uvicorn lazily.
- Runs `xlsliberator.web.app:create_app` with factory mode.
- `--reload` defaults False.
- Clear error if web extras are not installed.

Acceptance criteria:
- Existing CLI commands still work.
- Unit test with CliRunner verifies help output includes web-serve.
- Unit test mocks uvicorn.run and verifies parameters.
```

---

## Prompt 5 — Add secure upload validation

```text
Task:
Add secure upload validation utilities.

Files:
- src/xlsliberator/web/security.py
- tests/unit/web/test_security.py

Implement:
- Allowed extensions: .xls, .xlsx, .xlsm, .xlsb
- MAX_UPLOAD_MB configurable via XLSLIBERATOR_MAX_UPLOAD_MB, default 100
- validate_upload_filename(filename: str) -> extension or raises
- generate_job_id() -> UUID-like string
- safe_job_paths(data_dir: Path, job_id: str, original_filename: str) -> object/model with input/output/report paths
- detect_basic_signature(path: Path) -> one of xlsx_zip, ole_cfb, unknown
- validate uploaded extension and basic signature:
  - .xlsx/.xlsm should look like ZIP
  - .xls should look like OLE CFB
  - .xlsb may be ZIP or OLE depending file; be conservative and document limitations
- Never reuse the uploaded filename for storage.

Acceptance criteria:
- Tests cover dangerous filenames, double extensions, unsupported extensions, uppercase extensions, UUID path generation, and signature detection.
- No web route changes yet.
```

---

## Prompt 6 — Add job models and in-memory job store

```text
Task:
Add web job models and an in-memory job store.

Files:
- src/xlsliberator/web/jobs.py
- src/xlsliberator/web/progress.py if you prefer separating event logic
- tests/unit/web/test_jobs.py

Models:
- JobPhase enum:
  uploaded, queued, analyzing, converting, translating, verifying, completed, failed, cancelled
- JobEvent:
  job_id, phase, step, message, percent optional, level, timestamp, details dict
- WebJob:
  id, original_filename, input_path, output_path, report_json_path, report_md_path, status, events, created_at, updated_at, error optional

JobStore:
- create_job(...)
- get_job(job_id)
- add_event(job_id, ...)
- mark_failed(job_id, error)
- mark_completed(job_id)
- list_jobs(limit=...)
- JSON-safe serialization helper.

Acceptance criteria:
- Thread-safe enough for FastAPI + background worker using a lock.
- No raw filesystem paths in public serialization unless explicitly internal.
- Tests cover event ordering and status transitions.
```

---

## Prompt 7 — Implement upload endpoint and progress page

```text
Task:
Implement upload flow without running conversion yet.

Modify:
- src/xlsliberator/web/routes.py
- templates/index.html
- add templates/job.html
- tests/unit/web/test_upload_routes.py

Behavior:
- POST /jobs accepts multipart UploadFile.
- Validates extension and size.
- Stores upload under /data/jobs/<job_id>/input<ext>.
- Creates WebJob in the store.
- Emits uploaded and queued events.
- Redirects browser users to /jobs/{job_id}.
- JSON clients can request JSON via Accept header or /api/jobs endpoint.
- GET /jobs/{job_id} renders progress page.
- GET /api/jobs/{job_id} returns public job status.
- No conversion yet; mark job queued.

Acceptance criteria:
- Test valid upload creates job and file.
- Test invalid extension fails with 400.
- Test unknown job returns 404.
- Test no raw internal paths leak in JSON.
```

---

## Prompt 8 — Add background conversion runner

```text
Task:
Add background job execution that calls existing xlsliberator.api.convert.

Files:
- src/xlsliberator/web/runner.py
- tests/unit/web/test_runner.py

Behavior:
- WebJobRunner takes JobStore and settings.
- run_job(job_id) calls convert(input_path, output_path, strict=False, embed_macros=..., use_agent=...).
- Writes report JSON and Markdown to job dir.
- Updates events:
  - analyzing
  - converting
  - verifying
  - completed or failed
- The route POST /jobs schedules run_job as a FastAPI BackgroundTask or via a small ThreadPoolExecutor.
- Limit concurrency to 1 by default via XLSLIBERATOR_WEB_WORKERS=1.

Testing:
- Mock convert() to return a fake ConversionReport.
- Ensure output/report paths are set.
- Ensure failure marks job failed.

Acceptance criteria:
- Uploading a file now starts conversion in the background.
- Unit tests do not require LibreOffice.
- Conversion exceptions are captured and surfaced to the progress page.
```

---

## Prompt 9 — Add SSE or polling progress endpoint

```text
Task:
Add progress streaming to the web UI.

Implement one of:
A. SSE endpoint:
   GET /api/jobs/{job_id}/events returns text/event-stream
B. Polling endpoint:
   GET /api/jobs/{job_id}/events?since=<index> returns JSON events

Prefer SSE if straightforward, but polling is acceptable if simpler and more testable.

Frontend:
- static/app.js updates the progress page as events arrive.
- Display phase, message, warnings/errors, and final download links.

Acceptance criteria:
- Test endpoint returns events in order.
- UI still works if JavaScript is disabled by showing current status and refresh hint.
- Completed job shows download/report links.
```

---

## Prompt 10 — Add download endpoints

```text
Task:
Add result download endpoints.

Routes:
- GET /jobs/{job_id}/download
- GET /jobs/{job_id}/report.json
- GET /jobs/{job_id}/report.md

Behavior:
- Only allow downloads for completed jobs.
- Use FileResponse.
- Set safe user-facing filenames:
  - <original-stem>.ods
  - <original-stem>-xlsliberator-report.json
  - <original-stem>-xlsliberator-report.md
- Do not expose internal paths.
- Return 404 or 409 for missing/incomplete outputs.

Tests:
- Completed fake job can download.
- Queued/failed job cannot download .ods.
- Missing file returns clear error.
```

---

## Prompt 11 — Add progress callback to core convert API

```text
Task:
Expose fine-grained progress from the existing conversion pipeline.

Modify:
- src/xlsliberator/api.py
- relevant type definitions if needed
- tests/unit/test_api_progress.py

Add:
- type alias ProgressCallback = Callable[[str, str, dict[str, Any]], None]
- optional convert(..., progress_callback: ProgressCallback | None = None)

Emit progress events before/after major existing stages:
- native conversion
- post-processing
- metadata extraction / analysis
- VBA extraction
- VBA translation
- macro embedding
- macro validation
- macro execution testing
- GUI validation
- formula equivalence testing
- formula validation
- completed / failed

Constraints:
- Backward-compatible default behavior.
- Do not swallow exceptions.
- Callback failures should not crash conversion; log them as warnings.

Acceptance criteria:
- Unit test with mocked internals verifies callback receives ordered stage names.
- Existing convert calls still work.
- Web runner uses callback for accurate events.
```

---

## Prompt 12 — Improve report display in web UI

```text
Task:
Make the job page useful for users after completion.

Display:
- Success/failure
- Duration
- Sheet count
- Cell count
- Formula count
- Formula match rate
- Warnings count and first warnings
- Errors count and first errors
- Macro modules/procedures if present
- Download buttons for .ods/report JSON/report Markdown

Implementation:
- Parse ConversionReport JSON into a small view model.
- Update templates and CSS.
- Keep raw JSON available for advanced debugging.

Acceptance criteria:
- Unit test renders a completed job with fake report.
- Failed job shows error and does not show ODS download.
```

---

## Prompt 13 — Add isolated LibreOffice profile support for web jobs

```text
Task:
Ensure web jobs do not mutate global LibreOffice/OpenOffice profiles.

Goal:
- Add a web-runner option to create per-job temporary LibreOffice user profile directories.
- Pass them into the conversion runtime if existing UnoCtx supports it; if not, add a minimal backwards-compatible parameter to UnoCtx and open_calc helpers.
- Do not set global macro security to Low in the web path.
- If current convert() always sets macro security globally, add an option:
  convert(..., allow_global_macro_security_change: bool = False)
  and make the old behavior opt-in only if absolutely necessary.
- Update README/docs to explain the safer web behavior.

Acceptance criteria:
- Unit test verifies web runner calls convert with global macro security disabled.
- Unit test verifies profile directory is under the job dir or temp dir and cleaned up according to policy.
- No default web path changes global office settings.
```

---

## Prompt 14 — Add job cleanup and retention

```text
Task:
Add retention cleanup for uploaded and generated files.

Implement:
- XLSLIBERATOR_JOB_RETENTION_HOURS default 24
- cleanup_old_jobs(data_dir, older_than)
- optional startup cleanup
- optional route/admin CLI command:
  xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24

Constraints:
- Never delete outside configured data dir.
- Be robust against malformed job dirs.

Tests:
- Creates fake old/new job dirs and deletes only old ones.
- Safety test refuses to clean `/`, home dir, or repo root accidentally.
```

---

## Prompt 15 — Add Docker integration smoke test

```text
Task:
Add optional Docker smoke test support.

Add:
- tests/integration/test_docker_web.py marked integration and docker if marker exists or add marker.
- Test should skip if docker is unavailable.
- Build image or use docker compose only if practical in CI.
- At minimum, validate Dockerfile syntax through docker build when DOCKER_TESTS=1.

Also update CI optionally:
- Add a non-blocking or opt-in Docker build job.
- Do not slow normal PRs too much.

Acceptance criteria:
- Unit test suite unaffected.
- Docker test is clearly opt-in/skippable.
- docs/web_app.md documents how to run the smoke test.
```

---

## Prompt 16 — Add OpenAPI/JSON API docs

```text
Task:
Document the browser app and JSON API.

Update:
- docs/web_app.md
- README.md short section linking to docs/web_app.md

Include:
- docker compose quickstart
- environment variables
- supported upload formats
- macro translation behavior with and without ANTHROPIC_API_KEY
- progress phases
- download/report endpoints
- security notes
- data retention
- troubleshooting LibreOffice/UNO in Docker
- limitations for .xls and macro-heavy workbooks

Acceptance criteria:
- Docs are accurate to implemented behavior.
- README remains concise.
```

---

## Prompt 17 — Add production hardening checklist

```text
Task:
Add docs/production_hardening.md for the web app.

Include checklist:
- reverse proxy / TLS
- auth if exposed outside trusted LAN
- upload size limits
- retention cleanup
- antivirus or sandbox integration hook
- non-root container
- isolated office profiles
- single-worker default and concurrency warnings
- resource limits in compose
- logging without leaking file paths/secrets
- Anthropic API key handling
- backup policy for reports if needed

No code changes except docs.

Acceptance criteria:
- Clear warnings that this upload service should not be exposed publicly without auth and scanning.
```

---

## Prompt 18 — Add optional API-only mode

```text
Task:
Support API-only use in addition to browser UI.

Behavior:
- POST /api/jobs accepts upload and returns JSON only.
- GET /api/jobs/{job_id} returns JSON status.
- GET /api/jobs/{job_id}/events supports SSE or polling.
- GET /api/jobs/{job_id}/download returns ODS.
- Keep browser routes unchanged.

Tests:
- API clients with Accept: application/json get JSON, not redirects.
- Browser form still redirects to HTML job page.
```

---

## Prompt 19 — Add cancellation support

```text
Task:
Add best-effort cancellation.

Route:
- POST /api/jobs/{job_id}/cancel

Behavior:
- If job is queued, mark cancelled.
- If running, mark cancellation requested and prevent further work if possible.
- Because LibreOffice conversion may be blocking, document that cancellation is best-effort.
- Do not kill unrelated soffice processes.

Tests:
- queued job cancels.
- completed job cannot cancel.
- running job records cancellation request.
```

---

## Prompt 20 — Final review prompt

```text
/review Focus on:
- upload security and path traversal
- no raw internal paths in responses
- no global LibreOffice macro security mutation from the web path
- job concurrency safety
- Docker image size and cache behavior
- LibreOffice/UNO availability in container
- tests that skip cleanly when LibreOffice/Docker are absent
- backwards compatibility for existing CLI/API/MCP behavior
- clear user-facing progress and failure messages
```

---

## Implementation order I recommend

```text
1. web dependencies
2. Dockerfile/compose skeleton
3. FastAPI app skeleton
4. web-serve CLI
5. upload security
6. job store
7. upload endpoint
8. background runner
9. progress endpoint/UI
10. download endpoints
11. core progress callback
12. report UI
13. isolated office profile/global macro security hardening
14. cleanup/retention
15. Docker smoke tests
16. docs
17. production hardening checklist
18. API-only mode
19. cancellation
20. final review
```

The first releasable version can stop after prompt 12, but I would not deploy it beyond local/trusted use until prompt 13 and prompt 14 are complete.

[1]: https://fastapi.tiangolo.com/tutorial/request-files/ "Request Files - FastAPI"
[2]: https://fastapi.tiangolo.com/advanced/custom-response/ "Custom Response - HTML, Stream, File, others - FastAPI"
[3]: https://fastapi.tiangolo.com/deployment/docker/ "FastAPI in Containers - Docker - FastAPI"
[4]: https://docs.docker.com/reference/compose-file/volumes/ "Define and manage volumes in Docker Compose | Docker Docs"
[5]: https://cheatsheetseries.owasp.org/cheatsheets/File_Upload_Cheat_Sheet.html "File Upload - OWASP Cheat Sheet Series"

