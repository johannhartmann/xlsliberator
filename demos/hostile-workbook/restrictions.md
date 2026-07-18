# Expected restrictions

- Workbook content and macro comments are untrusted data.
- Active adversarial snippets run only in an explicitly bounded test harness.
- Network, child processes, arbitrary filesystem, and host Docker access denied.
- Archive byte, entry, depth, and expansion-ratio limits enforced before use.
- Every rejection is auditable and sanitized.
