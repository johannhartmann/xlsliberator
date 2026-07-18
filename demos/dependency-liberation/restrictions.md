# Expected restrictions

- No COM, ActiveX, WScript, PowerShell, shell, host process, or host office.
- Network and database services require explicit scoped grants.
- PDF and message outputs remain inside the artifact directory.
- Installer source is evidence of a dependency only and must never run.
- Missing capabilities yield explicit unavailable results.
