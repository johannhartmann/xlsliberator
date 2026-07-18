# Expected restrictions

- Run only in the pinned Docker LibreOffice runtime.
- Do not call Windows APIs, host processes, host devices, or arbitrary paths.
- Timer work must be cancellable and must not monopolize the UI thread.
- Sound is optional and must remain unavailable unless a portable capability is
  explicitly granted.
- The source workbook is immutable input.
