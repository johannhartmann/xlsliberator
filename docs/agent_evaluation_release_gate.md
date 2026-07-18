# Agent evaluation release gate

The deterministic package accepts an Open-SWE benchmark report only through
`xlsliberator.agent_evaluation`. The schema is the cross-repository contract
between the migration orchestrator and the package release workflow.

The report must:

- target LibreOffice full build `26.2.4.2`;
- include every one of the fourteen migration evaluators exactly once per case;
- preserve `passed`, `failed`, `skipped`, `unavailable`, and `not_run`;
- keep public and hidden aggregate summaries separate;
- group both partitions by approved team configuration, source format, and
  feature family;
- exclude all hidden definitions;
- bind every evaluator to a confined evidence path.

Required evaluator results must pass. A skipped generic-repair evaluator is
permitted only when it is explicitly non-required because no generic defect was
found. Any required corpus, security, hidden-acceptance, save/reopen, mutation,
dependency-removal, or reviewer approval failure keeps the case and report
non-release-ready.

`tools/write_release_inputs.py` validates the report before it can set
`agent_evaluation_passed=true`. `generate_capability_report` then includes the
`agent-evaluation` gate in the same fail-closed release decision as the corpus,
security, runtime identity, and evidence-schema gates.

The report contains evidence paths and aggregate reviewer outcomes. It never
contains hidden tests, hidden inputs, expected values, prompts, or private model
reasoning. Percentages are derived from observed passed/failed counts only; no
manually maintained success percentage is accepted as evidence.
