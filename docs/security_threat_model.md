# Untrusted workbook execution threat model

## Security objective

XLSLiberator treats every workbook, embedded formula, VBA comment, macro,
package part, scenario value, and agent-visible evidence string as attacker
controlled. The security objective is to prevent that data from reading or
changing arbitrary host files, reaching a network, inheriting credentials,
starting uncontrolled processes, changing tool policy, or manufacturing passing
evidence or certification.

An isolated LibreOffice user profile is useful for correctness but is **not** a
security boundary. LibreOffice, UNO and PyUNO execute only inside a disposable
Docker sandbox. If Docker or the immutable target image is unavailable, required
target, macro, and GUI execution is `UNAVAILABLE`; there is no host fallback.

## Trust boundaries

- The host orchestrator may resolve immutable image IDs, create a private job
  directory, copy explicitly allowed inputs, invoke Docker, and copy declared
  outputs back to an allowed workspace root.
- The web application container is also a trusted orchestrator and may receive
  the Docker socket. It translates only its dedicated runtime staging root to
  the corresponding host bind path. This control-plane container is outside the
  untrusted workbook boundary and must remain loopback-only or behind an
  authenticated gateway.
- Workbook inputs are constrained to roots configured with
  `XLSLIBERATOR_WORKSPACE_ROOTS` (or `--workspace-root` for MCP). Traversal,
  symlink escapes, special files, and output symlink replacement are rejected.
- Package inspection rejects unsafe member paths, ZIP symlinks, excessive
  compressed/uncompressed/part sizes, excessive compression ratios, entry
  floods, and oversized formula or macro text before Office runs.
- The container gets a read-only input mount and one isolated writable `/job`
  mount. Its root filesystem is read-only. `/tmp` and its disposable home are
  bounded tmpfs mounts.
- The container has no network, no inherited model/cloud credentials, no Docker
  socket, no added devices, no IPC namespace sharing, no Linux capabilities,
  and `no-new-privileges`. It runs as UID/GID 10001 with CPU, memory, PID,
  file-size, writable-space, and wall-time limits. `--init`, forced container
  removal on timeout, and `--rm` provide process-tree cleanup.
- Images are inspected before a job. Jobs run the resolved `sha256:` image ID,
  record both the configured reference and digest, and reject tag drift.
- `EnvironmentManifest` declares external capabilities. Only explicitly granted
  capabilities are passed to a job, and legacy plus typed grants are recorded in
  the evidence bundle and runtime trace.
- Coding-agent build and test commands use `DockerCommandSandbox`. Commands,
  images and mounts come from trusted application configuration, never workbook
  evidence or an agent proposal. The Docker socket is never mounted into those
  jobs.
- MCP binds to `127.0.0.1` by default. The Compose service may bind within its
  explicitly marked trusted container to a wildcard address only while Docker
  publishes that port on host `127.0.0.1`. Other non-loopback bindings are
  rejected. A remotely exposed deployment must add transport authentication and
  per-tool authorization rather than disabling this check.

## Prompt injection

Workbook text never becomes system or tool policy. Text sent to a coding model
is size bounded, hashed where applicable, and wrapped in an
`UNTRUSTED_WORKBOOK_DATA` or `UNTRUSTED_WORKBOOK_EVIDENCE` element. Agent output
is only a candidate patch. It cannot set certification; deterministic scenario,
trace-diff, build, and regression gates make that decision.

The malicious fixture set covers infinite macro loops, process spawning, host
file reads, network attempts, ZIP bombs, oversized formulas/macros, package path
traversal, ZIP symlinks, and prompt injection in VBA comments and cell text.

## Source oracle and higher assurance

Microsoft Excel source execution is not local and cannot use the LibreOffice
Docker image. It uses the same versioned `SandboxPolicy` contract through a
separately secured Windows remote worker or microVM and must return that sandbox
attestation in its runtime identity. Missing or invalid attestation fails closed.

Docker shares a host kernel. Deployments handling hostile documents with a
higher assurance requirement should implement the existing `microvm` or
`remote_worker` backend contract with ephemeral VMs, measured images, outbound
firewalling, short-lived worker identity, encrypted job transport, and remote
attestation. Evidence must state which backend actually ran; a remote or patched
runtime must never be labelled as the stock local Docker target.

## Residual risks

Docker engine and kernel vulnerabilities, parser bugs before sandbox entry, and
trusted-orchestrator compromise remain in scope for platform hardening. The
Windows worker must enforce its attested microVM/remote policy outside the Python
process. Public MCP exposure is unsupported until authenticated authorization is
configured. These limitations are explicit and cannot be converted into passing
certification evidence.
