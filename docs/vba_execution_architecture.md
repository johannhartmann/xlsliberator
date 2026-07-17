# Typed VBA execution architecture

XLSLiberator no longer treats module-by-module generated Python as a semantic
authority. VBA is first parsed into the versioned `VBAProjectIR` defined in
`src/xlsliberator/vba_ir.py`. Every module, declaration, procedure, parameter,
statement, expression, dependency, and conditional-compilation block carries a
stable source span and node ID. Raw source remains attached so a construct is
never invented when the deterministic parser cannot classify it.

## Execution decision

`build_execution_plan()` selects one strategy for every procedure:

1. use the native compatibility layer only when that capability is explicitly
   granted;
2. interpret only the statement subset implemented by `TypedVBAInterpreter`;
3. use a typed target compiler only when its runtime capability is present;
4. accept a Python translation only when a procedure-specific Microsoft Excel
   source trace and LibreOffice target trace are recorded as equivalent;
5. otherwise report the procedure as unavailable.

Missing COM, API, XLL, file, network, database, add-in, or project-reference
capabilities are typed in the environment manifest and block full
certification. Merely producing syntactically valid Python cannot make an
execution plan pass.

## Compatibility object model

`runtime/object_model.py` defines Application, Workbooks/Workbook,
Worksheets/Worksheet, Range/Cells, Names, worksheet tables and filters, charts,
pivots, controls, WorksheetFunction, calculation, and events. The object model
depends on the `CompatibilityBackend` protocol:

- `FakeExcelBackend` is a deterministic, pure in-memory backend for unit tests.
- `UnoExcelBackend` is the real target backend. It accepts an already-open UNO
  document and is instantiated only inside the guarded LibreOffice Docker
  worker. The module never imports host UNO.

The pinned office image copies the compatibility runtime into
`/opt/xlsliberator/src`; its bundled Python, UNO, and PyUNO identities are
checked during the Docker build. The Docker integration probe opens a real ODS,
uses the UNO backend for range access, collection enumeration, calculation, and
event dispatch, then closes without persisting the probe mutation.

## Source differential conformance

`vba_conformance.py` defines versioned micro-programs for coercion, ByRef,
default properties, arrays, error handling, classes, events, and range
operations. When licensed Windows Excel fixtures and the secured Windows oracle
are available, `generate_windows_micro_traces()` writes real source traces. A
fake oracle or missing fixture is reported as unavailable and never becomes
Excel evidence.

## Migration from the legacy translator

The provider-backed translator remains a candidate generator during migration.
It now parses the complete project first, embeds semantic node IDs in its source
map, emits the procedure execution plan in its evidence, and returns `partial`
when differential proof is absent. Existing call sites should migrate in this
order:

1. extract the full `VBAProjectIR`, including references and conditional build
   arguments;
2. declare and grant typed external capabilities in the scenario environment;
3. record Windows micro/source traces and matching Docker LibreOffice traces;
4. build the procedure execution plan;
5. embed only candidates whose chosen strategy is executable;
6. certify only after the exact macro scenario passes save/close/reopen with no
   source/target regression.

Unsupported parser constructs and object-model operations remain explicit
unavailable states. Expanding an LLM prompt is not an implementation strategy.
