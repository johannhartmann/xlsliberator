# XLSLiberator Web App User Guide

## What the Application Does

XLSLiberator converts Microsoft Excel workbooks into LibreOffice/OpenOffice Calc `.ods` files. The web app provides a browser-based workflow for uploading a workbook, tracking conversion progress, and downloading the converted spreadsheet plus conversion reports.

The current web app is intended for local or trusted internal use. It accepts spreadsheet files, stores them in server-managed job directories, runs the existing XLSLiberator conversion pipeline, and presents the results through a job page and JSON API.

Supported upload formats are:

- `.xlsx`
- `.xlsm`
- `.xls`
- `.xlsb`

The generated spreadsheet is an `.ods` file, which can be opened in LibreOffice Calc or OpenOffice Calc.

## Basic Workflow

1. Open the web app at `http://127.0.0.1:8080/`.
2. Choose an Excel workbook from the upload form.
3. Submit the workbook.
4. The app creates a conversion job and redirects to a job progress page.
5. The progress page updates as the workbook moves through upload, analysis, conversion, macro handling, verification, and completion.
6. When the job completes, download the converted `.ods` file, the JSON report, or the Markdown report.

If JavaScript is disabled, the progress page still shows the current job status and can be refreshed manually.

## Running the App with Docker

Build and run the Docker image:

```bash
docker build -t xlsliberator-web:test .
docker run --rm --name xlsliberator-web-preview -p 8080:8080 xlsliberator-web:test
```

Then open:

```text
http://127.0.0.1:8080/
```

The Docker image includes LibreOffice and the Python UNO bridge packages needed for conversion. Inside Docker, job artifacts are stored under `/data`.

You can also use Docker Compose:

```bash
docker compose up --build
```

## Upload Handling and File Safety

Uploaded files are treated as untrusted input. The app does not use the original filename as a storage path. Instead, it generates a UUID job ID and stores files under:

```text
/data/jobs/<job_id>/
```

For each job, the app writes server-controlled files such as:

- `input.<extension>` for the uploaded workbook
- `output.ods` for the converted spreadsheet
- `report.json` for machine-readable conversion details
- `report.md` for a human-readable summary
- `lo-profile/` for the job-specific LibreOffice profile

The app validates the uploaded filename, rejects unsupported extensions, blocks path traversal, checks for dangerous double extensions, and performs a basic container signature check. For example, `.xlsx` and `.xlsm` files must look like ZIP-based OOXML workbooks, while `.xls` files must look like OLE compound files.

The default upload limit is 100 MB. It can be changed with:

```bash
XLSLIBERATOR_MAX_UPLOAD_MB=50
```

## What Happens During Conversion

After upload, the job is added to an in-memory job store and submitted to a bounded background worker. By default, the app runs one conversion at a time because LibreOffice is resource-heavy and profile-sensitive.

The conversion runner calls the core `xlsliberator.api.convert()` pipeline. That pipeline can perform:

- Native LibreOffice conversion to `.ods`
- Workbook analysis
- Formula and named range repair
- VBA extraction
- VBA-to-Python macro translation, when configured
- Macro embedding
- Formula validation and equivalence checks
- Macro and GUI/event validation where available

For web jobs, XLSLiberator uses a per-job LibreOffice profile directory and disables global LibreOffice macro security mutation. This keeps the web flow from changing the user's normal LibreOffice profile settings.

## Progress States

The progress page displays structured job events. Common states include:

- `uploaded`: the server received and stored the file
- `queued`: the job is waiting for a worker
- `analyzing`: XLSLiberator is inspecting workbook structure
- `converting`: LibreOffice/native conversion and repair work is running
- `translating`: VBA extraction, translation, or macro embedding is running
- `verifying`: formulas, macros, or events are being checked
- `completed`: the `.ods` and reports are available
- `failed`: conversion stopped with an error
- `cancelled`: cancellation was requested before completion

Progress updates are exposed through:

```text
GET /api/jobs/{job_id}/events?since=0
```

The browser UI polls this endpoint and appends new events to the job page.

## Downloads and Reports

Completed jobs expose these download links:

```text
GET /jobs/{job_id}/download
GET /jobs/{job_id}/report.json
GET /jobs/{job_id}/report.md
```

The `.ods` download is the converted spreadsheet. The JSON report is intended for automation and debugging. The Markdown report is easier to read and share.

The report may include:

- Whether conversion succeeded
- Duration
- Sheet count
- Cell count
- Formula count
- Formula match rate
- Warning and error counts
- First warnings and errors
- Macro module and procedure counts
- Macro test results where available

Reports use safe user-facing filenames derived from the original workbook name, but internal server paths are not exposed.

## JSON API Usage

The web app can be used without the HTML interface.

Create a job:

```bash
curl -F "file=@workbook.xlsx" \
  -H "Accept: application/json" \
  http://127.0.0.1:8080/api/jobs
```

Check status:

```bash
curl http://127.0.0.1:8080/api/jobs/<job_id>
```

Poll events:

```bash
curl "http://127.0.0.1:8080/api/jobs/<job_id>/events?since=0"
```

Cancel a queued or running job:

```bash
curl -X POST http://127.0.0.1:8080/api/jobs/<job_id>/cancel
```

Cancellation is best-effort. A queued job can be cancelled before it starts. A running LibreOffice conversion may not stop immediately because some conversion steps are blocking.

## Configuration

The main environment variables are:

| Variable | Default | Purpose |
| --- | --- | --- |
| `XLSLIBERATOR_DATA_DIR` | `/data` | Root directory for jobs and artifacts |
| `XLSLIBERATOR_MAX_UPLOAD_MB` | `100` | Maximum accepted upload size |
| `XLSLIBERATOR_WEB_WORKERS` | `1` | Number of background conversion workers |
| `XLSLIBERATOR_JOB_RETENTION_HOURS` | `24` | Age after which job directories can be cleaned |
| `XLSLIBERATOR_EMBED_MACROS` | `1` | Whether translated macros should be embedded |
| `XLSLIBERATOR_USE_AGENT` | `1` | Whether agent-assisted macro translation is enabled |
| `ANTHROPIC_API_KEY` | unset | Optional key for higher-quality macro translation |

Without `ANTHROPIC_API_KEY`, macro-heavy workbooks may still convert, but macro translation quality and coverage may be limited.

## Cleanup and Retention

Old job directories are cleaned according to the configured retention window. The default retention period is 24 hours.

Manual cleanup:

```bash
xlsliberator cleanup-jobs --data-dir /data --older-than-hours 24
```

Cleanup refuses unsafe targets such as `/`, the user's home directory, or the repository root.

## Health Checks

The app exposes two health endpoints:

```text
GET /healthz
GET /readyz
```

`/healthz` confirms that the FastAPI process is responding. `/readyz` checks whether the data directory is writable and whether `soffice` or `libreoffice` is available.

## Limitations

Legacy `.xls` parsing is incomplete and should not be represented as fully validated. Macro-heavy workbooks may need manual review, especially when no LLM API key is configured. Some LibreOffice operations are blocking, so cancellation cannot always interrupt an active conversion immediately.

This upload service should not be exposed to the public internet without authentication, TLS, rate limiting, malware scanning, resource limits, and operational monitoring.
