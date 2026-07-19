# VBA execution architecture

XLSLiberator does not ship an embedded VBA translator, provider SDK, Excel
object-model emulator, repair agent, or alternative orchestrator. Open-SWE is
the only supported agent and orchestrator.

Open-SWE reads the extracted VBA and workbook dossier, generates target-native
Python/UNO modules, and submits those artifacts to deterministic XLSLiberator
primitives. XLSLiberator embeds the supplied modules, runs them in the pinned
LibreOffice `26.2.4.2` Docker runtime, and emits explicit evidence. It never
loads a model or decides which provider Open-SWE should use.

Direct UNO helpers should be added only for concrete recurring Calc operations.
An Excel compatibility runtime is not an accepted destination. If Open-SWE does
not supply target-native modules for source VBA, conversion reports the
unresolved capability instead of silently dropping or pretending to translate
the code.
