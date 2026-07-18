# Interactive game migration

The source is a real Excel VBA Tetris implementation with keyboard polling,
timed state changes, shapes, workbook/sheet events, score persistence, and
Windows sound calls. The migration must replace those surfaces with portable
LibreOffice controls, events, and bounded scheduling while preserving gameplay.

Start with `task.md`, then treat `acceptance.yaml` as the complete public
behavior contract. No target is supplied or implied.
