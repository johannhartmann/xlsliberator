# MCP Tool Contract

XLSLiberator's MCP server exposes LibreOffice as a Docker-only capability. The
host process may parse packages and orchestrate jobs; it never imports PyUNO or
starts a host office process.

Start the trusted local Docker orchestrator with:

```bash
mkdir -p artifacts/runtime-tmp artifacts/mcp-workspace
docker compose build libreoffice-runtime xlsliberator-mcp
docker compose up -d xlsliberator-mcp
```

The endpoint is `http://127.0.0.1:8000/mcp`. Files exposed to tools must be
placed below `artifacts/mcp-workspace`. The MCP control-plane container may
reach Docker; every workbook/LibreOffice worker it creates remains disposable,
networkless, and socket-free.

## Canonical response

Every tool response contains these fields independently:

| Field | Meaning |
|---|---|
| `transport_success` | The request/response boundary completed without timeout or protocol failure. |
| `operation_status` | `passed`, `failed`, `skipped`, `unavailable`, or `not_run`. Only `passed` means success. |
| `success` | Compatibility projection of `operation_status == passed`. |
| `implemented` | The requested operation has an implementation. |
| `capability_available` | The configured environment can currently provide it. |
| `evidence` | Structured inventories, reports, or worker responses supporting the result. |
| `error` | Typed error object, or `null`; never a friendly replacement for failure. |

Operation data such as `sheets`, `controls`, or `report` remains at the top
level for backward compatibility. The typed source of truth is
`xlsliberator.boundary_models.BoundaryResponse`.

## Runtime selection

Docker-backed tools accept `runtime_options` with these explicit fields:

```json
{
  "target": "libreoffice",
  "target_runtime_image": "xlsliberator-libreoffice:26.2.4.2",
  "target_runtime_digest": null,
  "timeout_seconds": 30,
  "resource_limits": {
    "network": "none",
    "read_only_root": true,
    "non_root": true,
    "cap_drop": ["ALL"],
    "no_new_privileges": true,
    "pids_limit": 256,
    "memory": "2g",
    "cpus": 2,
    "file_size_limit_blocks": 1048576
  },
  "workspace_root": null,
  "evidence_destination": null
}
```

The resource policy cannot be expanded by a caller. The image must equal the
configured image identity. When a digest is supplied, it must equal
`XLSLIBERATOR_LIBREOFFICE_IMAGE_ID`; arbitrary images and executable paths are
rejected. The worker resolves the configured reference to an immutable local
image ID and records that ID in evidence.

## Control actions

`list_controls` and `list_event_bindings` read discovered package inventories.
There are no hardcoded button names.

- `execute_button_handler` resolves an inventoried control's script URI and
  invokes that handler directly in the disposable target container.
- `click_form_button` is a separate capability and currently returns
  `implemented=false`, `capability_available=false`, and `UNAVAILABLE`.
- `send_keyboard_input`, `open_document_gui`, and `take_screenshot` are likewise
  explicit unavailable capabilities. They never claim actions occurred.

## Runtime validation

`validate_document_runtime` requires every open, recalculate, save, close,
reopen, and package stage to pass and verifies that the staged source was not
mutated. `agent_validator` consumes this complete evidence together with macro,
control, and event inventories. Reading sample cells such as A1-C1 cannot pass
agent validation.
