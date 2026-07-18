---
name: windows-dependency-replacement
description: Use this skill when workbook behavior depends on Windows, COM, Office automation, databases, HTTP, XLLs, or proprietary add-ins needing open replacements.
compatibility: Docker-only XLSLiberator; replacements are target-native open services with explicit grants and LibreOffice 26.2.4.2 integration.
recommended-tools: read_file write_file xlsprobe migration-check libreoffice-runtime-mcp
---

# Windows dependency replacement

Liberate behavior through explicit open capabilities. Provider implementations
belong in adapters, never in the workbook migration core.

## Use when

Use when forensics finds COM automation, Windows APIs, Access/ADODB, network
clients, filesystem automation, XLL/UDFs, or proprietary Office services. Do not
use to silently remove required behavior or grant unrestricted network/host access.

## Replacement map

- Outlook → SMTP or open mail API behind a `mail` capability.
- ADODB/Access → DB-API, JDBC, SQLite, PostgreSQL, or LibreOffice Base.
- FileSystemObject → `pathlib` or UNO file service with scoped export roots.
- WinHTTP/XMLHTTP → open HTTP client with host/method allowlist.
- `user32`/`kernel32` → UNO/AWT or a Linux-native scoped service.
- XLL/UDF → Python/UNO add-in or LibreOffice extension.
- Word automation → Writer UNO.

## Inputs and outputs

Inputs: call sites, data contracts, authentication needs, errors, side effects,
deployment constraints, and user-approved grants. Outputs: dependency inventory,
target-native service interfaces, adapters and deterministic mocks, capability
manifest, tests, operational configuration, and evidence of legacy removal.

## Tool sequence

1. Trace every dependency from source entry point through inputs, outputs,
   side effects, retries, and error behavior.
2. Distinguish required behavior from provider-specific mechanics.
3. Select an open replacement and define a narrow typed capability contract.
4. Implement core logic against the contract and providers in adapters.
5. Test with deterministic mocks, denial, timeout, malformed result, idempotency,
   and a real authorized integration when available.
6. Record grants and verify the sandbox remains network-denied otherwise.
7. Prove COM, Office, Windows DLL, Excel runtime, and unresolved add-in absence.

## Failure handling

Missing credentials or authorized capability is `UNAVAILABLE`, not success.
Never embed credentials, run provider CLIs in core code, or preserve a
proprietary runtime as a hidden fallback.

## Acceptance checklist

- [ ] Each proprietary dependency has an approved behavior-level replacement.
- [ ] Capabilities are least-privilege, recorded, mockable, and revocable.
- [ ] Provider-specific code is isolated in adapters.
- [ ] Error and side-effect behavior is source-derived and tested.
- [ ] Legacy runtime dependencies are absent from output.

## Tested examples

Positive: replace Access queries with bundled SQLite fixtures through a DB-API
repository, test query results and unavailable database behavior, then verify
the dashboard in LibreOffice.

Adversarial: copy an Access database to the sandbox and start a Windows VM to run
ADODB. Reject it as a retained proprietary runtime.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
