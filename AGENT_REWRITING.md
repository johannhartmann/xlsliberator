# Agent-Based VBA Rewriting System

## Overview

The agent-based rewriting system uses semantic analysis and multi-phase orchestration to translate complex VBA code (especially games and event-driven applications) into properly architected Python-UNO code.

### Key Features

- **Semantic Analysis**: Uses LLM to understand code intent, not regex pattern matching
- **Architecture Transformation**: Converts VBA patterns to Python-UNO equivalents
  - Blocking loops (`While...Wend` + `Sleep`) → Timer-based updates (`threading.Timer`)
  - Keyboard polling (`GetAsyncKeyState`) → Event listeners (`XKeyListener`)
  - Global state → Class-based state management
- **Iterative Refinement**: Validates and fixes generated code automatically
- **Knowledge-Based**: Uses comprehensive templates and examples for transformations

## When to Use Agent Rewriting

Use `--agent-rewrite` when your VBA code contains:

- **Games** (Tetris, Snake, Pong, etc.)
- **Keyboard input handling** (GetAsyncKeyState, KeyDown events)
- **Animation loops** (Sleep + DoEvents in loops)
- **Real-time updates** (game loops, timers)
- **Windows API calls** (user32.dll, kernel32.dll)
- **Event-driven architecture**

For simple VBA (basic formulas, button clicks, simple loops), the standard translation is faster and sufficient.

## Usage

```bash
# Convert with agent-based rewriting (experimental)
xlsliberator convert input.xlsm output.ods --agent-rewrite

# With strict mode (fail on any errors)
xlsliberator convert input.xlsm output.ods --agent-rewrite --strict

# Generate detailed report
xlsliberator convert input.xlsm output.ods --agent-rewrite --report report.json
```

### Requirements

- `ANTHROPIC_API_KEY` environment variable must be set
- Requires Claude Sonnet 4 model access

## Architecture

### Phase 1: Pattern Detection (`pattern_detector.py`)

Analyzes VBA modules to detect complexity patterns semantically:

- **Input**: VBA modules with source code, procedures, dependencies, API calls
- **Process**: LLM analyzes code structure and intent
- **Output**: Complexity level (simple/game/advanced/untranslatable) with confidence score

**Example Detection**:
```
Input: Tetris.xlsm with 5 VBA modules
Output: Complexity = "game" (95% confidence)
Detected patterns:
  - Windows API keyboard polling (GetAsyncKeyState)
  - Blocking game loop (While...DoEvents)
  - Real-time animation (Sleep timing)
```

### Phase 2: Architecture Design (`agent_rewriter.py`)

Designs the transformation strategy:

- **Input**: VBA modules + complexity analysis
- **Process**: LLM designs Python-UNO architecture using template knowledge
- **Output**: Structured design document with:
  - Transformation strategy
  - Required components (classes, listeners, etc.)
  - UNO services needed
  - Key pattern mappings
  - Implementation plan
  - Critical notes

**Example Design**:
```
Strategy: "Event-driven game loop with XKeyListener"
Components:
  - GameController (timer-based loop)
  - GameKeyListener (keyboard events)
  - Tetromino (piece class)
  - GameState (enum)
UNO Services:
  - com.sun.star.awt.XKeyListener
  - com.sun.star.awt.Key
Transformations:
  - GetAsyncKeyState() → XKeyListener.keyPressed()
  - While...DoEvents → threading.Timer
  - Sleep(500) → timer.interval = 0.5
```

### Phase 3: Code Generation (`agent_rewriter.py`)

Generates Python-UNO code per module:

- **Input**: VBA modules + architecture design
- **Process**: LLM generates Python-UNO code following architecture
- **Output**: Complete Python modules with:
  - Proper imports (uno, unohelper, threading, etc.)
  - Class-based architecture
  - Event handlers
  - g_exportedScripts tuple

**Example Generated Code**:
```python
import uno
import unohelper
from com.sun.star.awt import XKeyListener, Key
import threading
from enum import IntEnum

class GameState(IntEnum):
    STOPPED = 0
    RUNNING = 1
    PAUSED = 2

class GameController:
    def __init__(self, doc):
        self.doc = doc
        self.state = GameState.STOPPED
        self.timer = None
        self.frame_interval = 0.5

    def _schedule_next_frame(self):
        if self.state == GameState.RUNNING:
            self.timer = threading.Timer(self.frame_interval, self._update_frame)
            self.timer.daemon = True
            self.timer.start()

    # ...

g_exportedScripts = (start_game, stop_game, pause_game)
```

### Phase 4: Testing & Refinement (`agent_rewriter.py`)

Validates and iteratively improves generated code:

- **Input**: Generated code + output ODS path
- **Process**:
  1. Embed Python modules into ODS
  2. Validate syntax and structure
  3. If errors found, use LLM to generate fixes
  4. Re-embed and re-validate
  5. Repeat up to max_iterations (default: 5)
- **Output**: Validation result with:
  - Syntax validity
  - Export script presence
  - Error list
  - Warning list
  - Iterations used

## Templates and Knowledge Base

### `templates/game_architecture.py`

Comprehensive guide for VBA game transformations:

- **Problem description**: Why VBA patterns don't work in Python-UNO
- **Solution patterns**: Complete before/after examples
- **Code samples**: Working XKeyListener, timer loop, state management
- **Best practices**: Architecture checklist, testing strategy, common pitfalls

### `templates/keyboard_listener.py`

XKeyListener template with:

- Complete class implementation
- Key mappings (arrows, space, letters, etc.)
- Registration helper function
- Pattern variations (game movement, WASD, general control)

### `templates/timer_loop.py`

Timer-based loop template with:

- State machine (stopped/running/paused)
- Frame scheduling with threading.Timer
- Update/render separation
- Button handler functions
- Pattern variations (falling piece, continuous movement, turn-based, animation)

## Known Limitations

### Performance

- **Sequential module generation**: Each module takes ~1 minute to generate
- **Total time**: ~5-10 minutes for a 5-module project like Tetris
- **Cause**: Synchronous API calls to Claude for each module
- **Future optimization**: Batch generation or parallel API calls

### Scope

- **Games only**: Currently optimized for game-like VBA (keyboard input, loops, animation)
- **Windows API**: Focused on user32.dll (keyboard) and kernel32.dll (Sleep, GetTickCount)
- **Simple VBA**: Overkill for basic macros - use standard translation instead

### Validation

- **Syntax only**: Validates Python syntax and g_exportedScripts presence
- **No execution testing**: Games require user interaction, can't be tested automatically
- **Manual verification needed**: Open the ODS and test the game manually

## Testing Results

### Tetris.xlsm Test

**Input**:
- 5 VBA modules (ThisWorkbook, Sheet1, Tetromino, Game, Engine)
- 36 procedures total
- 60 Windows API calls (GetAsyncKeyState, Sleep, GetTickCount, Beep)
- Complex game loop with keyboard polling

**Results**:
- ✅ Phase 1: Correctly detected "game" complexity (95% confidence)
- ✅ Phase 2: Designed 8-component architecture (XKeyListener, timer-based loop, state machine)
- ✅ Phase 3: Began generating event-driven Python-UNO code
- ⏱️ Phase 4: Did not complete due to performance (timeout after 5 minutes)

**Key Achievements**:
- Semantic analysis correctly identified game patterns
- Architecture design included proper UNO services (XKeyListener, threading.Timer)
- Generated code showed correct structure (class-based, event-driven)
- No regex pattern matching used (as required)

## Future Improvements

### Short Term

1. **Batch generation**: Generate all modules in one API call
2. **Caching**: Cache architecture designs for similar projects
3. **Progress UI**: Show real-time progress during long conversions

### Medium Term

1. **Parallel execution**: Generate multiple modules concurrently
2. **Partial results**: Save intermediate results on timeout
3. **Resume capability**: Continue from last successful phase
4. **Execution testing**: Test simple functions (not games) automatically

### Long Term

1. **Fine-tuned model**: Train on VBA→Python-UNO examples
2. **Incremental generation**: Generate and validate incrementally
3. **User feedback loop**: Learn from manual corrections
4. **Pattern library**: Build reusable transformation patterns

## API Reference

### AgentRewriter Class

```python
from xlsliberator.agent_rewriter import AgentRewriter

agent = AgentRewriter()
generated_code, validation = agent.rewrite_vba_project(
    modules=vba_modules,
    source_file="input.xlsm",
    output_path=Path("output.ods"),
    max_iterations=5
)
```

### CLI

```bash
xlsliberator convert INPUT OUTPUT --agent-rewrite [OPTIONS]

Options:
  --agent-rewrite     Use multi-agent system for complex VBA rewriting
  --strict           Fail on any errors
  --report PATH      Save detailed report (JSON or Markdown)
```

### Return Values

#### GeneratedCode

```python
@dataclass
class GeneratedCode:
    modules: dict[str, str]           # module_name → Python code
    architecture_doc: str             # Architecture design reference
    completeness_score: float         # 0.0-1.0 confidence
    known_limitations: list[str]      # Known issues
```

#### ValidationResult

```python
@dataclass
class ValidationResult:
    syntax_valid: bool                # Python syntax OK
    has_exports: bool                 # g_exportedScripts present
    execution_successful: bool        # Execution test passed
    errors: list[str]                 # Error messages
    warnings: list[str]               # Warning messages
    iterations_used: int              # Refinement iterations
```

## Troubleshooting

### "ANTHROPIC_API_KEY required"

Set your API key:
```bash
export ANTHROPIC_API_KEY="sk-ant-..."
```

### "Agent-based rewriting failed"

- Check API key is valid
- Check network connectivity
- Try with `--strict` to see full error details
- Check logs for specific error messages

### Timeout during conversion

- This is expected for large projects (5+ modules)
- The system is working, just slow (~1 min per module)
- Future versions will optimize this
- For now, consider:
  - Converting smaller files first
  - Using standard translation for simple VBA
  - Waiting longer (set higher timeout)

### Generated code doesn't work

- Manual testing required for games
- Open the ODS in LibreOffice
- Test button clicks and keyboard input
- Check LibreOffice macro security is set to "Low"
- Review generated code in Scripts/python/ folder
- Compare with template examples in `templates/`

## Examples

See `tests/data/Tetris.xlsm` for a complete example of VBA code suitable for agent rewriting.

Generated output: `tests/output/Tetris_agent.ods` (after full conversion completes)

## Contributing

To improve the agent rewriting system:

1. **Add templates**: Create new pattern templates in `templates/`
2. **Improve prompts**: Refine system prompts in `agent_rewriter.py` and `pattern_detector.py`
3. **Optimize performance**: Implement batch generation or caching
4. **Add patterns**: Extend detection to new VBA patterns (COM automation, database access, etc.)

## License

Same as xlsliberator project.
