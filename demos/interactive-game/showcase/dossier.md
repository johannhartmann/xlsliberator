# Interactive game source dossier

This is a source-derived migration dossier for the `interactive-game` episode.
It inventories the immutable Excel input and the public acceptance contract. It
does **not** claim that a migration ran, that a target exists, that LibreOffice
accepted an output, that any scenario passed, or that a reviewer approved a
result.

## Immutable source

- Workbook: `demos/interactive-game/source/TetrisGameDemo.xlsb`
- Format: Microsoft Excel binary workbook (`.xlsb`)
- Size: 69,074 bytes
- SHA-256: `da1bddc2c20ed8f5557b547e04a84cb1b476eca010e30a6be549be650894e4d1`
- Upstream project: `zq99/tetris-game-excel-vba`
- Upstream repository: <https://github.com/zq99/tetris-game-excel-vba>
- Upstream commit: `83d3c4ab9b5bc8629aa0f520ca4ea668a6ff05eb`
- Source license: MIT, copyright 2020 zq99

The workbook and supplied upstream modules and assets are immutable inputs.
Migration work must use a copy and must never rewrite these files. The workbook
contains `xl/vbaProject.bin`; that is evidence about the source, not an
executable dependency that may be retained in a migrated target.

## Supplied modules

| Module | Kind | Source behavior |
| --- | --- | --- |
| `ModGame.bas` | standard VBA module | Game state, piece generation and rotation, keyboard polling, blocking timing loop, collision handling, line collapse, scoring, preview rendering, and screen controls |
| `ModHighScore.bas` | standard VBA module | Username lookup, high-score qualification, Windows-path random-access file persistence, score-sheet population, and sorting |
| `Sheet1.cls` | `Game` worksheet class | Selection-change guard for `Score`, `totalRows`, and `$D$26` |
| `Sheet2.cls` | `Score` worksheet class | Activate and selection-change guards for the high-score table |
| `ThisWorkbook.cls` | workbook class | Open-time game-sheet selection and before-close saved-state mutation |

The complete procedure and file-hash inventory is in
`source-inventory.json`.

## Workbook controls and events

Static package and exported-source inspection identifies these source controls:

| Control | Sheet | Excel macro binding | Source role |
| --- | --- | --- | --- |
| `cmdStartStop` | `Game` | `Run` | Toggle the game between Start and Stop |
| `cmdAbout` | `Game` | `About_Game` | Display game information |
| `cmdHighScores` | `Game` | `ViewHighScores` | Open the high-score sheet |
| `cmdReturn` | `Score` | `ViewGameScreen` | Return to the game sheet |

The source also declares `Workbook_Open`, `Workbook_BeforeClose`,
`Worksheet_Activate`, and two `Worksheet_SelectionChange` handlers. Keyboard
behavior is polled through `GetAsyncKeyState`: left and right move the active
piece, down performs a faster drop, and up or Ctrl rotates it. The workbook
instructions describe Escape as pause, while the exported code implements
cancellation through the loop's error path. The target contract therefore
requires explicit, observable start, pause, resume, and reset transitions
rather than copying exception-driven cancellation.

## State and scoring

- The visible playfield is `M4:AA31`, with named left, right, top, and bottom
  boundaries.
- Named state includes `Score`, `totalRows`, `previewblock`, `HighScores`,
  `ScoreData`, and `ScoreColumn`.
- The source defines nine piece variants and tracks the active cells, collision
  base points, orientation, next piece, landed cells, game state, and timer
  delay in module variables.
- A fast downward move adds 1 point, a landed piece adds 5 points, and a
  completed row adds 100 points and increments `totalRows` by one.
- The delay starts at 160 milliseconds and decreases in row-count tiers to 40
  milliseconds.

These are source observations. Public acceptance expectations remain
authoritative; cached workbook values and source quirks are not an Excel oracle.

## Proprietary and host dependencies

| Source dependency | Source usage | Required disposition |
| --- | --- | --- |
| `user32.dll!GetAsyncKeyState` | Keyboard polling | Remove; replace with a declared LibreOffice-native input surface |
| `kernel32!Sleep` | Blocking game-loop delay | Remove; replace with bounded, cancellable scheduling that does not block the LibreOffice UI thread |
| `winmm.dll!PlaySoundA` | Landing and line-collapse audio | Unresolved/optional; no sound may run unless a portable capability is explicitly granted |
| Excel Shapes/VML macro bindings | Buttons and handlers | Replace with target-native controls and non-Basic event bindings |
| Excel Range/Union/ColorIndex APIs | Board and piece state | Replace with direct Python/UNO or target-native document operations |
| `Environ("username")` and `Application.UserName` | High-score identity | Remove ambient host identity dependency or replace it with an explicit bounded value |
| `ExceltrisHighScores.txt` beside the workbook | Random-access score persistence using a Windows path | Replace with document-contained or job-confined persistence |

The supplied `land.wav` and `collapse.wav` files are source evidence only.
Sound remains unresolved and optional. No host sound device, host process,
arbitrary host path, Windows API, COM server, Excel runtime, VBA runtime,
LibreOffice Basic program, or proprietary add-in is granted by this dossier.

## Public acceptance sources

- Task: `demos/interactive-game/task.md`
- Acceptance: `demos/interactive-game/acceptance.yaml`
- Restrictions: `demos/interactive-game/restrictions.md`
- Provenance: `demos/interactive-game/PROVENANCE.md`

The public contract defines five scenarios:

1. `keyboard-control`: left, right, down, and rotate produce one valid state
   transition each, while a wall-crossing input leaves the board unchanged.
2. `timer-tick`: three eligible ticks advance the piece and a control event is
   processed before the third tick without freezing the UI.
3. `native-controls`: start, pause, resume, and reset produce their documented
   transitions; reset clears score and board without reopening.
4. `document-events`: open initializes controls once, then a move, save, close,
   and reopen preserve game state and high score.
5. `line-collapse`: the public near-complete board collapses exactly once and
   applies the public scoring rule.

All five require execution in the pinned LibreOffice `26.2.4.2` Docker runtime
before any pass can be claimed.

## Evidence still required

This dossier leaves all target and certification claims open. Completion
requires, at minimum, a runnable target and digest; generated Python/UNO and
native-control inventories; content-addressed open, recalculate, interaction,
event, save, close, reopen, and assertion traces; public scenario results;
mutation results; privacy-safe hidden-test results; proprietary-dependency
removal evidence; an independent reviewer verdict; and a replayable demo.
Missing, skipped, unavailable, not-run, partial, or inconclusive evidence must
not be represented as success.
