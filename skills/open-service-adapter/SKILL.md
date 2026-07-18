---
name: open-service-adapter
description: Use this skill when migrated workbook behavior needs a reusable target-native external service with explicit capabilities, providers, and deterministic mocks.
compatibility: Docker-only XLSLiberator; adapters integrate with LibreOffice 26.2.4.2 and receive only server-authorized scoped capabilities.
recommended-tools: read_file write_file migration-check libreoffice-runtime-mcp
---

# Open service adapter

Create a narrow, target-native service boundary for behavior that should not
live inside the document.

## Use when

Use for reusable mail, database, HTTP, filesystem export, document generation,
or similar external behavior after dependency planning. Do not use for native
UNO operations, provider lock-in, unrestricted network access, or hiding missing
functionality behind a mock.

## Inputs and outputs

Inputs: source behavior contract, approved open provider options, capability
grant, data classification, retry/idempotency semantics, errors, and deployment
constraints. Outputs: typed interface, provider adapters, deterministic mock,
configuration schema, authorization checks, audit metadata, unit/integration
tests, and runtime evidence.

## Tool sequence

1. Derive the smallest behavior-level contract from source and requirements.
2. Define typed requests/results, timeouts, errors, idempotency, and bounded logs.
3. Require an explicit capability and restrict hosts, methods, paths, identities,
   and data as applicable.
4. Implement provider-neutral core interface, open provider adapter(s), and a
   deterministic offline mock.
5. Inject the adapter into direct Python/UNO logic; never expose credentials to
   the document or sandbox.
6. Test success, denial, timeout, malformed response, retry, duplicate request,
   and redaction.
7. Run an authorized target interaction and record grant plus service evidence.

## Failure handling

No grant, provider, or required secret yields `UNAVAILABLE`; mocks cannot satisfy
a required production integration gate. Avoid automatic permission expansion
and never log payloads that may contain workbook/customer data.

## Acceptance checklist

- [ ] Contract represents behavior rather than one provider SDK.
- [ ] Core code is provider-neutral and mockable.
- [ ] Capability and authorization are explicit and least-privilege.
- [ ] Secrets remain server-side and logs are bounded/redacted.
- [ ] Failure/idempotency behavior and real target integration are tested.

## Tested examples

Positive: a PDF-mail adapter accepts a generated artifact reference and message
metadata, uses a scoped SMTP provider, and has an offline mock for acceptance.

Adversarial: import a cloud vendor SDK directly in translated workbook logic and
read its token from environment variables. Reject it.

## Global anti-patterns

No Excel worker, VBA runtime, `ExcelContext` expansion, provider-specific core
code, or success without target execution.
