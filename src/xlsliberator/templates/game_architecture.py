"""Architecture transformation guide for game applications.

This module provides architectural knowledge for transforming VBA game code
(with keyboard input, game loops, etc.) into LibreOffice Python-UNO equivalents.
"""

GAME_ARCHITECTURE_GUIDE = """
# Architecture Transformation Guide: VBA Games → LibreOffice Python-UNO

## Core Problem

VBA games typically use:
1. **GetAsyncKeyState()** - Polls keyboard state synchronously
2. **While/Do loops** - Blocking game loop that polls input
3. **Sleep()** - Blocking delays for frame timing
4. **DoEvents** - Yields to allow UI updates

LibreOffice Python-UNO requires:
1. **Event-driven architecture** - XKeyListener for keyboard events
2. **Timer-based updates** - threading.Timer for non-blocking frame updates
3. **Async state management** - Non-blocking state handling

## Transformation Strategy

### 1. Keyboard Input Transformation

**VBA Pattern:**
```vba
Public Sub GiveControlls(tetr As Object)
    If GetAsyncKeyState(vbKeyDown) Then
        Call tetr.Move("down")
    ElseIf GetAsyncKeyState(vbKeyUp) Then
        Call tetr.Rotate("right")
    ElseIf GetAsyncKeyState(vbKeyLeft) Then
        Call tetr.Move("left")
    ElseIf GetAsyncKeyState(vbKeyRight) Then
        Call tetr.Move("right")
    End If
End Sub
```

**Python-UNO Pattern:**
```python
import uno
import unohelper
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import Key

class GameKeyListener(unohelper.Base, XKeyListener):
    \"\"\"Keyboard event listener for game input.\"\"\"

    def __init__(self, game_controller):
        self.game = game_controller
        self.pressed_keys = set()

    def keyPressed(self, event):
        \"\"\"Handle key press events.\"\"\"
        self.pressed_keys.add(event.KeyCode)

        # Handle input immediately
        if event.KeyCode == Key.DOWN:
            if self.game.current_piece:
                self.game.current_piece.move("down")
        elif event.KeyCode == Key.UP:
            if self.game.current_piece:
                self.game.current_piece.rotate("right")
        elif event.KeyCode == Key.LEFT:
            if self.game.current_piece:
                self.game.current_piece.move("left")
        elif event.KeyCode == Key.RIGHT:
            if self.game.current_piece:
                self.game.current_piece.move("right")
        elif event.KeyCode == Key.SPACE:
            if self.game.current_piece:
                self.game.current_piece.move("drop")

    def keyReleased(self, event):
        \"\"\"Handle key release events.\"\"\"
        self.pressed_keys.discard(event.KeyCode)

    def disposing(self, event):
        \"\"\"Clean up on disposal.\"\"\"
        pass

# Register listener with document
def register_keyboard_listener(doc, game_controller):
    \"\"\"Register keyboard listener on document window.\"\"\"
    listener = GameKeyListener(game_controller)

    # Get window and add key handler
    window = doc.getCurrentController().getFrame().getContainerWindow()
    window.addKeyListener(listener)

    return listener
```

### 2. Game Loop Transformation

**VBA Pattern:**
```vba
Sub Game()
    While GameState <> Stopped
        ' Update game state
        GiveControlls(tetr)
        tetr.Move("down")

        ' Draw frame
        DrawScreen()

        ' Frame timing
        Sleep(500)
        DoEvents
    Wend
End Sub
```

**Python-UNO Pattern:**
```python
import threading

class GameController:
    \"\"\"Main game controller with timer-based updates.\"\"\"

    def __init__(self, doc):
        self.doc = doc
        self.game_state = GameState.STOPPED
        self.current_piece = None
        self.timer = None
        self.frame_interval = 0.5  # 500ms

    def start_game(self):
        \"\"\"Start the game loop.\"\"\"
        self.game_state = GameState.RUNNING
        self.current_piece = Tetromino()
        self._schedule_next_frame()

    def stop_game(self):
        \"\"\"Stop the game loop.\"\"\"
        self.game_state = GameState.STOPPED
        if self.timer:
            self.timer.cancel()
            self.timer = None

    def pause_game(self):
        \"\"\"Pause the game.\"\"\"
        self.game_state = GameState.PAUSED

    def resume_game(self):
        \"\"\"Resume the game.\"\"\"
        self.game_state = GameState.RUNNING
        self._schedule_next_frame()

    def _schedule_next_frame(self):
        \"\"\"Schedule the next frame update.\"\"\"
        if self.game_state == GameState.RUNNING:
            self.timer = threading.Timer(self.frame_interval, self._update_frame)
            self.timer.daemon = True  # Allow program to exit
            self.timer.start()

    def _update_frame(self):
        \"\"\"Update one game frame.\"\"\"
        if self.game_state != GameState.RUNNING:
            return

        # Automatic piece movement
        if self.current_piece and self.current_piece.can_move("down"):
            self.current_piece.move("down")
        else:
            # Piece landed - check for completed rows, spawn new piece
            self._handle_piece_landed()

        # Render
        self._draw_screen()

        # Check game over
        if self._is_game_over():
            self.stop_game()
            return

        # Schedule next frame
        self._schedule_next_frame()

    def _draw_screen(self):
        \"\"\"Render the game state to spreadsheet.\"\"\"
        # Get active sheet
        sheet = self.doc.getSheets().getByIndex(0)

        # Render game board cells...
        # (existing rendering code)
```

### 3. State Management Transformation

**VBA Pattern:**
```vba
' Global state variables
Public GameState As Integer
Public score As Long
Public tetr As Object
```

**Python-UNO Pattern:**
```python
from enum import IntEnum

class GameState(IntEnum):
    \"\"\"Game state enumeration.\"\"\"
    STOPPED = 0
    RUNNING = 1
    PAUSED = 2

class GameModel:
    \"\"\"Game state management.\"\"\"

    def __init__(self):
        self.state = GameState.STOPPED
        self.score = 0
        self.current_piece = None
        self.next_piece = None
        self.game_board = [[0 for _ in range(10)] for _ in range(20)]

    def reset(self):
        \"\"\"Reset game state.\"\"\"
        self.state = GameState.STOPPED
        self.score = 0
        self.current_piece = None
        self.next_piece = None
        self.game_board = [[0 for _ in range(10)] for _ in range(20)]
```

### 4. Integration Pattern

**Complete Game Integration:**
```python
import uno
from loguru import logger

# Global game controller instance
_game_controller = None

def init_game(*args):
    \"\"\"Initialize game (called from button or on document load).\"\"\"
    global _game_controller

    doc = XSCRIPTCONTEXT.getDocument()

    if _game_controller is None:
        # Create game components
        model = GameModel()
        controller = GameController(doc, model)

        # Register keyboard listener
        listener = register_keyboard_listener(doc, controller)

        _game_controller = controller
        logger.info("Game initialized")

    return _game_controller

def start_game(*args):
    \"\"\"Start/resume game (button handler).\"\"\"
    controller = init_game()

    if controller.game_state == GameState.STOPPED:
        controller.start_game()
    elif controller.game_state == GameState.PAUSED:
        controller.resume_game()
    elif controller.game_state == GameState.RUNNING:
        controller.pause_game()

def reset_game(*args):
    \"\"\"Reset game (button handler).\"\"\"
    controller = init_game()
    controller.stop_game()
    controller.model.reset()

# Export for LibreOffice
g_exportedScripts = (init_game, start_game, reset_game)
```

## Key Architecture Changes

### VBA → Python-UNO Transformations:

1. **Synchronous polling → Asynchronous events**
   - `GetAsyncKeyState()` → `XKeyListener.keyPressed()`
   - Immediate response, no polling needed

2. **Blocking loop → Timer-based updates**
   - `While...Wend` + `Sleep()` → `threading.Timer`
   - Non-blocking, allows UI responsiveness

3. **Global state → Class-based state**
   - Module-level variables → Class instances
   - Better encapsulation and lifecycle management

4. **Direct button access → Event handlers**
   - `Buttons("StartButton").Caption = "Pause"` → Button event binding
   - Use button click events, not direct property manipulation

5. **DoEvents → Not needed**
   - Timer-based updates are naturally non-blocking
   - UI remains responsive

## Implementation Checklist

When transforming a VBA game:

- [ ] Identify all keyboard input points → Create XKeyListener
- [ ] Identify main game loop → Create timer-based update method
- [ ] Identify global state → Create state management classes
- [ ] Identify button handlers → Create exported functions
- [ ] Remove all Windows API calls (GetAsyncKeyState, Sleep)
- [ ] Replace synchronous loops with timer scheduling
- [ ] Test keyboard responsiveness
- [ ] Test frame timing accuracy
- [ ] Test pause/resume functionality
- [ ] Test game state persistence

## Common Pitfalls to Avoid

1. **Don't** try to keep the synchronous loop structure
2. **Don't** poll for keyboard state in a timer callback
3. **Don't** use `time.sleep()` in event handlers or timers
4. **Don't** create new listener instances repeatedly
5. **Don't** forget to cancel timers on game stop
6. **Don't** access document from non-main thread without synchronization

## Testing Strategy

1. **Unit test** each component (Tetromino class, collision detection, etc.)
2. **Integration test** keyboard listener registration
3. **Integration test** timer-based frame updates
4. **Manual test** actual gameplay
5. **Performance test** frame rate consistency
"""

# Key UNO imports needed for games
UNO_GAME_IMPORTS = """
import uno
import unohelper
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import Key
from com.sun.star.lang import XEventListener
import threading
from enum import IntEnum
from loguru import logger
"""

# Key concepts the agent must understand
KEY_CONCEPTS = {
    "keyboard_input": "Use XKeyListener event-based pattern, not polling",
    "game_loop": "Use threading.Timer for non-blocking frame updates",
    "state_management": "Use class-based state, not module globals",
    "button_handlers": "Export functions via g_exportedScripts tuple",
    "frame_timing": "Use Timer.daemon=True and proper cancellation",
    "document_access": "Always use XSCRIPTCONTEXT.getDocument()",
}
