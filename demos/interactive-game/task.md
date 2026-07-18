# Task

Migrate `source/TetrisGameDemo.xlsb` to LibreOffice `26.2.4.2`. Preserve
keyboard movement, rotation, falling-piece timing, start/pause/reset controls,
document events, board state, scoring, high-score persistence, and save/reopen
behavior. Replace `GetAsyncKeyState`, `Sleep`, Windows sound APIs, and Excel
shape/event assumptions with bounded Python/UNO or native LibreOffice services.

Produce a migration dossier, runnable target, behavioral evidence for every
public scenario, and explicit dispositions for anything not supported.
