# XLSLiberator corpus layout

- `atomic/`: small feature fixtures and deterministic recipes.
- `composed/`: multi-feature minimized cases.
- `regressions/`: promoted generic failures with normalized signatures.
- `security/`: safe adversarial cases and denial expectations.
- `manifests/`: episode catalog, search index, and CI subsets.
- `public-scenarios/`: public acceptance index; hidden tests are absent.

`corpus/manifest.json` remains the low-level conformance index. The serious
episode catalog is `manifests/episodes.json`; together they drive evidence
reports without converting missing results into passes.
