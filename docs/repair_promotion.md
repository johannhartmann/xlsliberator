# Generic repair promotion

Reusable failures follow one evidence-bound path:

1. reproduce the exact behavior;
2. minimize it without changing the failure signature;
3. add a failing regression;
4. patch the classified owner layer;
5. rerun the exact scenario;
6. run the affected corpus;
7. obtain independent review;
8. open the focused upstream review.

`repairs/<repair-id>/record.json` binds every stage and artifact by SHA-256.
`RepairRecord.verify()` rejects missing files, identity drift, absent stock
failure, absent patched success, source-commit mismatch, missing affected-corpus
evidence, and non-independent review.

The checked-in `tdf-172479-text-functions` record is the real reference flow.
The pinned LibreOffice `26.2.4.2` source build fails the minimized TEXTAFTER
scenario, the same commit with the single declared patch passes, and the
runtime/Calc binary identities and upstream Gerrit review are preserved.

## MCP services

Run only inside the repository application image:

```bash
docker compose run --rm --service-ports test \
  xlsliberator corpus-mcp-serve --host 127.0.0.1 --port 8010

docker compose run --rm --service-ports test \
  xlsliberator buildfarm-mcp-serve --host 127.0.0.1 --port 8020
```

The corpus service exposes public search, prior-failure search, recorded public
suite validation, minimized-failure registration, run comparison, capability
reporting, and a reviewer-only hidden-result operation. Hidden definitions are
never returned.

The build-farm contract exposes source-worktree, patch, build, upstream-test,
artifact, comparison, and log operations. Mutation is disabled unless the
server is explicitly enabled and the Open-SWE role allowlist authorizes a
LibreOffice engineer. This repository does not silently fall back to a local
build: without the external isolated backend, mutation returns `UNAVAILABLE`.
