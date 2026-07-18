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

## Session API

The public MCP registry is deliberately closed:

| Lifecycle | Document and runtime |
|---|---|
| `create_session` | `open_document` |
| `collect_logs` | `inspect_document` |
| `destroy_session` | `list_sheets`, `read_cells`, `write_cells` |
|  | `list_formulas`, `recalculate`, `list_controls` |
|  | `dispatch_control_event`, `send_keyboard_event` |
|  | `execute_python_macro`, `capture_screenshot`, `export_pdf` |
|  | `save`, `close`, `reopen` |

`create_session` returns `session_id`. Every other tool rejects a missing,
unknown, destroyed, or mismatched session ID. A session records the exact
LibreOffice 26.2.4.2 image/executable identity, a unique profile identifier,
UNO port and display, a source-preserving working copy, and a durable evidence
directory. Failed creation and failed operations are archived so
`collect_logs` remains available after cleanup.

All input and output paths are resolved through configured workspace roots.
Inputs are copied into the session; the original is never opened by Office and
is hash-checked by scenario execution. Mutations replace only the session
working copy after a successful worker response. `save` and `export_pdf` copy
declared outputs atomically to allowed destinations.

## Runtime and status semantics

The caller cannot select an arbitrary image, executable, port, display, Docker
policy, or host fallback. The service resolves the configured pinned image to
an immutable ID, probes its exact executable and bundled PyUNO, and records the
identity in the session descriptor. Each Office invocation uses the session's
unique profile identity and connection resources in a disposable, networkless,
read-only-root container. Wall-time expiry force-removes the named container
before returning `UNAVAILABLE`.

The MCP transport status and workbook operation status are independent. A
well-formed response can truthfully report `transport_success=true` and
`operation_status=unavailable`; only `operation_status=passed` sets
`success=true`.

## UI actions and deprecated names

`dispatch_control_event`, `send_keyboard_event`, and `capture_screenshot` must
use and prove a real UI/event layer. The pinned headless runtime does not
currently provide that layer, so these operations return
`operation_status=unavailable`, `capability_available=false`, and a typed
`UIEventLayerUnavailable` error. Direct script-handler invocation is not a
control click.

The loose legacy tool names, including `execute_button_handler`,
`click_form_button`, `send_keyboard_input`, `open_document_gui`, and
`take_screenshot`, are not registered by the public server. The old
`xlsliberator mcp-serve` command remains only as a deprecated alias for
`xlsliberator libreoffice-mcp-serve`.

## Trust mode

The command binds to `127.0.0.1` by default and explicitly runs in
trusted-local mode. Non-loopback binding is rejected unless the Compose
control-plane container is marked as the trusted proxy and Docker publishes
the endpoint only on host loopback. Remote exposure requires authenticated
transport and per-tool authorization.
