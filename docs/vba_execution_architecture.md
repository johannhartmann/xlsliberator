# Legacy VBA execution architecture

The former typed VBA interpreter, Excel-shaped object model, provider clients,
and self-repair loop are not part of the deterministic XLSLiberator core.
They are retained temporarily under `xlsliberator.legacy_agent` behind the
optional `legacy-agent` dependency extra and emit a deprecation warning.

New migrations must not depend on that compatibility runtime. Open-SWE agents
read the original VBA and dossier directly and generate target-native Python/UNO
code. XLSLiberator then embeds that code, runs it in the pinned LibreOffice
Docker runtime, and emits deterministic evidence. Direct UNO helpers should be
added only for concrete, recurring target operations; an Excel object-model
emulator is not an accepted destination.

The prohibited Microsoft Excel oracle, Windows worker, VBA conformance runtime,
and source-execution documentation were removed. Acceptance evidence comes from
declared scenarios, independent review, hidden tests, mutation tests, and the
LibreOffice target runtime. Missing evidence remains unavailable or failed and
never becomes success.
