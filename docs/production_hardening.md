# Web App Production Hardening

Do not expose the upload service publicly without these controls:

- Put the app behind an authenticated reverse proxy with TLS and rate limits.
- Map each tenant to a stable, non-user-controlled
  `XLSLIBERATOR_OPEN_SWE_OWNER_ID`.
- Store `XLSLIBERATOR_OPEN_SWE_TOKEN` in a secret manager and rotate it.
- Permit web traffic only to the internal `xlsliberator-open-swe` service.
- Keep the web container unprivileged and never mount a Docker socket into it.
- Enforce upload limits at both proxy and app; scan files before migration.
- Keep `XLSLIBERATOR_WEB_WORKERS` bounded to protect Open-SWE.
- Configure CPU, memory, PID, and request-duration limits.
- Keep local job cleanup and Open-SWE owner/retention enforcement enabled.
- Verify source/dependency deletion before a migration is marked complete.
- Avoid logging bearer tokens, workbook content, internal paths, prompts,
  hidden tests, or model reasoning.
- Expose only the fixed, owner-checked artifact allowlist with `nosniff`.
- Back up deliverables only under an explicit retention and access policy.

LibreOffice `26.2.4.2`, its bundled Python, UNO, and PyUNO execute only in the
pinned office image controlled by Open-SWE. The web app must fail closed when
Open-SWE is unconfigured or unavailable; there is no local conversion fallback.
