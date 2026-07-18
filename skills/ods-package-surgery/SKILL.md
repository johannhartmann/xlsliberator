---
name: ods-package-surgery
description: Use this skill when a valid ODS package requires a narrow transactional XML, manifest, script, or event-binding modification with preservation evidence.
compatibility: Docker-only XLSLiberator; odstool validates packages consumed by LibreOffice 26.2.4.2.
recommended-tools: read_file write_file odstool migration-check libreoffice-runtime-mcp
---

# ODS package surgery

Use `odstool` transactions for package-level changes. Never edit a live output
archive in place.

## Use when

Use for targeted ODF XML, manifest, embedded Python, dialog, or event-binding
changes that cannot be made safely through UNO. Do not use for broad workbook
generation, source extraction, or hiding a LibreOffice import defect.

## Inputs and outputs

Inputs: immutable ODS, exact requested change, package invariants, and acceptance
cases. Outputs: a new ODS, transaction manifest/diff, verification log, preserved
entry hashes, and LibreOffice execution evidence.

## Tool sequence

1. Hash the input and run `odstool verify`.
2. Start a transaction in a fresh job directory and enumerate all ZIP entries,
   media types, namespaces, manifests, scripts, and event bindings.
3. Apply the smallest declared operation; use XML-aware edits, safe paths, and
   deterministic serialization.
4. Validate XML, mimetype placement/storage, manifest consistency, duplicate
   names, path normalization, signatures, and binding targets.
5. Compare untouched entry hashes and semantic inventories.
6. Commit atomically to a new file; retain the original.
7. Open, execute relevant interactions, save, close, reopen, and rerun
   `odstool verify` plus acceptance scenarios.

## Failure handling

Rollback on any failed operation or invariant. Reject traversal, absolute paths,
symlinks, duplicate normalized entries, oversized expansions, malformed XML,
or missing binding targets. If signed content must change, report signature
impact explicitly rather than silently stripping it.

## Acceptance checklist

- [ ] Input remains unchanged and output is atomically produced.
- [ ] Every mutation is declared in the transaction manifest.
- [ ] Unrelated entries are byte-identical or a justified canonicalization is recorded.
- [ ] Manifest, scripts, controls, and event bindings resolve.
- [ ] Package verification and real LibreOffice save/reopen pass.

## Examples

Positive: add one Python script and bind a button event; verify the script media
entry, binding URI, untouched sheet XML hashes, real button dispatch, and reopen.

Adversarial: a supplied archive contains `../../home/user/.ssh/config`. Reject it
before extraction and preserve the security failure evidence.
