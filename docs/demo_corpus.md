# Serious migration demo corpus

The serious corpus complements low-level fixtures with eight complete
migration episodes:

| Episode | Source | Primary surfaces | PR | Nightly |
|---|---|---|---|---|
| interactive game | XLSB | keyboard, timer, controls, events, state | no | yes |
| invoice workflow | XLSM | UserForms, events, PDF, Outlook replacement | yes | yes |
| database dashboard | XLSX + source bundle | database refresh and parameterized queries | no | yes |
| add-in replacement | XLS + XLL source | UDF and extension packaging | no | yes |
| legacy application | BIFF8 XLS | OLE/VBA inventory and reopen | yes | yes |
| operations model | complex XLSB | chart, comments, drawings, VBA | no | yes |
| dependency liberation | XLSM + source bundle | COM, paths, process, database | no | yes |
| hostile workbook | inert XLSX + safe adversarial source | denial and isolation | yes | yes |

The source catalog is
`tests/corpus/manifests/episodes.json`. `search-index.json` is a generated,
non-confidential corpus-MCP index; `subsets.json` defines PR, nightly, and
security sets. `public-scenarios/index.json` points to the complete public
acceptance contracts. Hidden tests are not stored in this repository.

## Truthful target status

No known-good target is checked in. Every episode target is `not_verified`
until a pinned LibreOffice `26.2.4.2` Docker run produces independently
reviewed evidence. Missing, skipped, unavailable, unsupported, waived, failed,
and passed results remain separate.

Generate the current evidence-derived feature/format report in Docker:

```bash
docker compose run --rm test xlsliberator demo-corpus-report \
  --results tests/corpus/manifests/not-run-results.json
```

Validate layout, source hashes, licenses, acceptance behavior, and search
metadata:

```bash
docker compose run --rm test xlsliberator demo-corpus-validate
```

The checked-in not-run result contains no fabricated execution. As real
episode evidence arrives, replace it with result records bound to artifact
paths and normalized failure signatures. Any generic failure must be minimized
into `tests/corpus/regressions/` before it is considered repaired.
