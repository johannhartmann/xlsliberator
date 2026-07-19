# XLSLiberator

[![CI](https://github.com/johannhartmann/xlsliberator/workflows/CI/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/ci.yml)
[![Security](https://github.com/johannhartmann/xlsliberator/workflows/Security/badge.svg)](https://github.com/johannhartmann/xlsliberator/actions/workflows/security.yml)
[![License: GPL v3](https://img.shields.io/badge/license-GPLv3-blue.svg)](LICENSE)

XLSLiberator is an Open-SWE agent for migrating Excel workbooks to LibreOffice
Calc. It accepts `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` sources and produces a
LibreOffice `.ods` deliverable with migration and review evidence.

Open-SWE is the only agent and orchestrator in this repository. LibreOffice
`26.2.4.2` is the only target. Docker is the only supported development and
runtime platform.

## How it works

For each workbook, Open-SWE:

1. inspects workbook structure, formulas, VBA, controls, and dependencies;
2. creates a baseline ODS through the pinned LibreOffice image;
3. delegates focused analysis to its workbook, VBA, and review specialists;
4. builds target-native repairs when the source requires them;
5. validates the result and performs a real save, close, and reopen cycle;
6. delivers the ODS and evidence only after independent review.

The repository's CLI, MCP methods, package editors, and validation modules are
tools used by Open-SWE. They are not a second migration workflow.

There is no separate XLSLiberator repository, maintained Open-SWE fork,
deterministic orchestrator, host-office fallback, or local PyUNO fallback.

## Quick start

Prerequisites:

- Docker with Docker Compose
- an explicitly selected Open-SWE model
- the credential for that model's provider

```bash
git clone https://github.com/johannhartmann/xlsliberator.git
cd xlsliberator
cp .env.example .env
```

Set `XLSLIBERATOR_OPEN_SWE_MODEL` and only the matching provider credential in
`.env`, then start the stack:

```bash
mkdir -p artifacts/runtime-tmp artifacts/open-swe-workspaces
docker compose up -d --build xlsliberator-web
```

Open `http://127.0.0.1:8080/`, upload a workbook or select one of the two small
web samples, and start the migration.

If no model is configured, the services remain healthy but migration creation
fails before a model run starts. GitHub Models is disabled by default and is
never selected as a fallback. It requires all three of:

```dotenv
XLSLIBERATOR_OPEN_SWE_MODEL=github_models:<model-id>
XLSLIBERATOR_GITHUB_MODELS_ENABLED=1
GITHUB_MODELS_TOKEN=<dedicated-token>
```

## Runtime boundary

The Compose stack contains three application services:

- `xlsliberator-web`: stores uploads, creates authenticated Open-SWE threads,
  streams progress, and serves owner-checked artifacts;
- `xlsliberator-open-swe`: runs the pinned upstream Open-SWE package with the
  XLSLiberator graph from this repository;
- `xlsliberator-mcp`: exposes the curated workbook and LibreOffice tools used by
  that graph.

Only the MCP application boundary receives the Docker socket so it can create
disposable office workers. The web and Open-SWE containers receive neither the
socket nor a host shell. LibreOffice, its bundled Python, UNO, PyUNO, and
`soffice` run only in the pinned office image.

See [Open-SWE architecture](docs/architecture/open-swe-migration.md) and
[web application](docs/web_app.md) for the detailed trust boundaries.

## Component tools

The following commands are useful for development and diagnosis. Run them only
through Docker; they do not replace the Open-SWE workflow.

```bash
# Source inventory
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator inspect "$PWD/input.xlsm"

# Baseline conversion
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator convert "$PWD/input.xlsx" "$PWD/output.ods"

# Validate an existing candidate
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator validate "$PWD/input.xlsx" "$PWD/output.ods"

# Full low-level transform and certification
docker compose --profile ci-runner run --rm test-runner \
  xlsliberator transform-validated "$PWD/input.xlsm" "$PWD/output.ods"
```

### Workbook forensics

`xlsprobe` creates a bounded, source-derived dossier without executing workbook
content:

```bash
mkdir -p artifacts/source-case
docker compose --profile ci-runner run --rm test-runner \
  xlsprobe dossier "$PWD/input.xlsm" --output "$PWD/artifacts/source-case"
```

Focused commands include `package-tree`, `extract-vba`, `formulas`, `controls`,
`dependencies`, and `previews`.

### ODS package editing

`odstool` performs transactional script and event-binding changes:

```bash
docker compose --profile ci-runner run --rm test-runner \
  odstool verify "$PWD/output.ods"
docker compose --profile ci-runner run --rm test-runner \
  odstool upsert-script "$PWD/output.ods" "$PWD/repair.py" --dry-run
```

Mutation commands verify the complete candidate before atomically replacing the
original package.

### LibreOffice MCP

The internal MCP service is started as part of the web stack. For explicit
component work:

```bash
mkdir -p artifacts/runtime-tmp artifacts/open-swe-workspaces
docker compose up -d --build xlsliberator-mcp
```

It exposes workbook inspection, baseline conversion, candidate building,
validation, and stateful LibreOffice session operations. See
[MCP tools](docs/mcp_tools.md).

## Configuration

Important Open-SWE variables are documented in `.env.example`:

- `XLSLIBERATOR_OPEN_SWE_MODEL`
- the selected provider credential
- `XLSLIBERATOR_OPEN_SWE_REASONING_EFFORT`
- `XLSLIBERATOR_OPEN_SWE_MAX_OUTPUT_TOKENS`
- `XLSLIBERATOR_OPEN_SWE_SERVICE_TOKEN`
- `XLSLIBERATOR_GITHUB_MODELS_ENABLED`

Model-free workbook and LibreOffice tools do not read provider credentials.

## Development

The host shell may run Docker, Git, and file operations only. Do not run host
Python, `uv`, PyUNO, UNO, LibreOffice, or `soffice`, including for diagnostics.

```bash
docker compose build test
docker compose run --rm test ruff format --check .
docker compose run --rm test ruff check src tests
docker compose run --rm test mypy src
docker compose run --rm test pytest tests/unit
make test-integration
make test-docker-web
```

`make all` runs the complete local CI sequence in Docker.

## Repository layout

```text
src/xlsliberator/
├── open_swe_agent/    Open-SWE graph, specialists, state, and curated tools
├── web/               FastAPI web client and browser UI
├── mcp_server.py      Curated internal MCP server
├── libreoffice_mcp.py Stateful LibreOffice session tools
├── docker_runtime.py  Disposable office-container boundary
├── lo_worker.py       Code executed inside the office image
├── xlsprobe.py        Bounded source inspection
├── odstool.py         Transactional ODS editing
└── validation_*.py    Transformation validation

docker/                Pinned application, Open-SWE, test, and office images
office/                LibreOffice source identity and maintained target patch
tests/                 Unit, Docker integration, and real-workbook tests
docs/                  Current architecture, runtime, and security documentation
```

## Limitations

- Migration quality depends on the specific workbook and configured model.
- Proprietary COM automation, external DLLs, and Windows-only dependencies must
  be replaced with target-native or open-service implementations.
- A baseline conversion is not proof of behavioral equivalence.
- Missing, skipped, unavailable, or not-run validation is never treated as
  success.

## License

XLSLiberator is licensed under the
[GNU General Public License v3.0 or later](LICENSE).
