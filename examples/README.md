# XLSLiberator examples

These examples demonstrate deterministic XLSLiberator tools. Open-SWE is the
only supported agent and orchestrator; there is no embedded provider or agent
SDK example.

## Docker setup

```bash
mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
docker compose build libreoffice-runtime xlsliberator-mcp
docker compose up -d xlsliberator-mcp
```

The host must not run Python, `uv`, LibreOffice, UNO, PyUNO, or `soffice`.

## Deterministic conversion

Run Python only inside the application container:

```bash
docker compose run --rm test python - <<'PY'
from pathlib import Path

from xlsliberator.api import convert

report = convert(
    input_path=Path("input.xlsx"),
    output_path=Path("output.ods"),
)
print(report.model_dump_json(indent=2))
PY
```

Source VBA is extracted but is not translated by this deterministic API.
Open-SWE must generate target-native Python/UNO modules and hand them to the
deterministic artifact-upsert and validation tools.

## Web flow

```bash
cp .env.example .env
# Set XLSLIBERATOR_OPEN_SWE_URL and XLSLIBERATOR_OPEN_SWE_TOKEN.
docker compose build xlsliberator-web
docker compose up -d xlsliberator-web
```

Open `http://127.0.0.1:8080/`. The web container delegates exclusively to
Open-SWE and fails closed if it is unavailable.

## See also

- [Web application](../docs/web_app.md)
- [MCP server](../docs/mcp_server.md)
- [API reference](../docs/api.md)
- [Open-SWE architecture](../docs/architecture/open-swe-migration.md)
