# Bounded agent repair

`xlsliberator agent-repair <evidence-bundle>` creates a versioned `AgentRun` for
a failing evidence bundle. The implementation deliberately separates proposal
generation from acceptance. Deterministic rules run first; a configured coding
agent may only propose a strict `CandidatePatch` when the rules do not resolve
the evidence. The proposal schema has no certification field.

Every candidate is applied to a fresh detached Git worktree. A repository lease
prevents concurrent repair runs from sharing a checkout. The acceptance path is
fixed and cannot be reordered by workbook content or an agent:

1. build the bounded scope;
2. execute the exact source scenario;
3. execute the exact LibreOffice target scenario;
4. compare the source and target traces;
5. run the focused regression subset.

Any failed, missing, reordered, or out-of-scope gate rejects the candidate. The
complete attempt and reason are atomically persisted before the worktree is
removed. Wall time, iterations, model cost, disk use, and build scope are
bounded. An accepted patch is preserved as `accepted.patch`; applying or
committing that patch is a separate repository-owner action.

## Trust boundary

Workbook-derived evidence is untrusted data. When supplied to a coding agent it
is size-bounded, hashed, and enclosed in an
`UNTRUSTED_WORKBOOK_EVIDENCE` element. It cannot provide system policy, tool
permissions, executable commands, deterministic gate results, or certification.
Only application-supplied gate callbacks may execute builds and tests.

The CLI supports `--dry-run`, which validates and persists the run without
creating a worktree. The default CLI has no coding-agent or executable gate
configuration and therefore fails closed for non-dry runs unless a pre-reviewed
`candidate.patch` exists; even that patch cannot pass without trusted gates.
LibreOffice, UNO, and PyUNO execution is always delegated to the pinned Docker
target runner and must never occur in the host repair process.

## Certification provenance

`repair_provenance_from_run` converts an accepted run into a hash-bound
`RepairProvenance` record. The record lists the deterministic gates and can be
rendered in a certification report, but `ValidationCertification` still derives
its `certified` value solely from required validation gates.
