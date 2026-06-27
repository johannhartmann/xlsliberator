# Web App Production Hardening

Do not expose the upload service publicly without additional controls.

Checklist:

- Put the app behind a reverse proxy with TLS.
- Require authentication outside a trusted LAN.
- Enforce upload size limits at the proxy and app.
- Keep `XLSLIBERATOR_WEB_WORKERS=1` unless LibreOffice isolation is load-tested.
- Configure CPU and memory limits in Compose or the orchestrator.
- Keep per-job LibreOffice profiles enabled.
- Never enable global macro security mutation for the web path.
- Add antivirus or sandbox scanning before conversion.
- Keep job retention cleanup enabled and verify it cannot delete outside `/data`.
- Run the container as a non-root user.
- Store `ANTHROPIC_API_KEY` in a secret manager, not Compose files or logs.
- Avoid logging raw filesystem paths, secrets, or workbook contents.
- Decide whether generated reports need backup or should expire with uploads.
