"""Timer-based game loop template for non-blocking updates."""

TIMER_LOOP_TEMPLATE = """
# Timer-Based Game Loop for LibreOffice Python-UNO

import threading
from enum import IntEnum
from loguru import logger


class {{STATE_ENUM_NAME}}(IntEnum):
    \"\"\"{{APPLICATION_NAME}} state enumeration.\"\"\"
    STOPPED = 0
    RUNNING = 1
    PAUSED = 2


class {{CONTROLLER_CLASS_NAME}}:
    \"\"\"Main game/application controller with timer-based updates.

    Manages application state, timing, and frame updates without blocking.
    \"\"\"

    def __init__(self, doc{{ADDITIONAL_PARAMS}}):
        \"\"\"Initialize controller.

        Args:
            doc: LibreOffice document instance
            {{ADDITIONAL_PARAMS_DESC}}
        \"\"\"
        self.doc = doc
        self.state = {{STATE_ENUM_NAME}}.STOPPED
        self.timer = None
        self.frame_interval = {{FRAME_INTERVAL}}  # seconds
        {{INIT_STATE_VARS}}

        logger.info("{{CONTROLLER_CLASS_NAME}} initialized")

    def start(self):
        \"\"\"Start the application/game.\"\"\"
        if self.state != {{STATE_ENUM_NAME}}.STOPPED:
            logger.warning("Already running or paused")
            return

        logger.info("Starting {{APPLICATION_NAME}}")
        self.state = {{STATE_ENUM_NAME}}.RUNNING
        {{ON_START_CODE}}
        self._schedule_next_frame()

    def stop(self):
        \"\"\"Stop the application/game.\"\"\"
        if self.state == {{STATE_ENUM_NAME}}.STOPPED:
            return

        logger.info("Stopping {{APPLICATION_NAME}}")
        self.state = {{STATE_ENUM_NAME}}.STOPPED

        # Cancel timer if running
        if self.timer:
            self.timer.cancel()
            self.timer = None

        {{ON_STOP_CODE}}

    def pause(self):
        \"\"\"Pause the application.\"\"\"
        if self.state != {{STATE_ENUM_NAME}}.RUNNING:
            return

        logger.info("Pausing {{APPLICATION_NAME}}")
        self.state = {{STATE_ENUM_NAME}}.PAUSED

        # Cancel timer
        if self.timer:
            self.timer.cancel()
            self.timer = None

    def resume(self):
        \"\"\"Resume from paused state.\"\"\"
        if self.state != {{STATE_ENUM_NAME}}.PAUSED:
            return

        logger.info("Resuming {{APPLICATION_NAME}}")
        self.state = {{STATE_ENUM_NAME}}.RUNNING
        self._schedule_next_frame()

    def toggle_pause(self):
        \"\"\"Toggle between running and paused states.\"\"\"
        if self.state == {{STATE_ENUM_NAME}}.RUNNING:
            self.pause()
        elif self.state == {{STATE_ENUM_NAME}}.PAUSED:
            self.resume()

    def _schedule_next_frame(self):
        \"\"\"Schedule the next frame update.\"\"\"
        if self.state == {{STATE_ENUM_NAME}}.RUNNING:
            self.timer = threading.Timer(self.frame_interval, self._update_frame)
            self.timer.daemon = True  # Allow program to exit cleanly
            self.timer.start()

    def _update_frame(self):
        \"\"\"Update one frame (called by timer).

        This is the main game/application update loop.
        \"\"\"
        # Early exit if stopped or paused
        if self.state != {{STATE_ENUM_NAME}}.RUNNING:
            return

        try:
            # Update game logic
            {{UPDATE_LOGIC}}

            # Render/draw
            {{RENDER_LOGIC}}

            # Check end conditions
            if {{END_CONDITION}}:
                self.stop()
                return

        except Exception as e:
            logger.error(f"Error in frame update: {e}")
            self.stop()
            return

        # Schedule next frame
        self._schedule_next_frame()

    def set_frame_rate(self, fps: float):
        \"\"\"Set frame rate (frames per second).

        Args:
            fps: Desired frames per second
        \"\"\"
        self.frame_interval = 1.0 / fps
        logger.debug(f"Frame rate set to {fps} FPS (interval: {self.frame_interval:.3f}s)")

    {{ADDITIONAL_METHODS}}


# Button handler functions (exported to LibreOffice)

_controller_instance = None


def init_{{APPLICATION_NAME_LOWER}}(*args):
    \"\"\"Initialize the application (called on first use).\"\"\"
    global _controller_instance

    if _controller_instance is None:
        doc = XSCRIPTCONTEXT.getDocument()
        _controller_instance = {{CONTROLLER_CLASS_NAME}}(doc{{ADDITIONAL_ARGS}})
        logger.info("{{APPLICATION_NAME}} controller created")

    return _controller_instance


def start_{{APPLICATION_NAME_LOWER}}(*args):
    \"\"\"Start the application (button handler).\"\"\"
    controller = init_{{APPLICATION_NAME_LOWER}}()

    if controller.state == {{STATE_ENUM_NAME}}.STOPPED:
        controller.start()
    elif controller.state == {{STATE_ENUM_NAME}}.PAUSED:
        controller.resume()
    elif controller.state == {{STATE_ENUM_NAME}}.RUNNING:
        controller.pause()


def stop_{{APPLICATION_NAME_LOWER}}(*args):
    \"\"\"Stop/reset the application (button handler).\"\"\"
    controller = init_{{APPLICATION_NAME_LOWER}}()
    controller.stop()


def pause_{{APPLICATION_NAME_LOWER}}(*args):
    \"\"\"Pause the application (button handler).\"\"\"
    controller = init_{{APPLICATION_NAME_LOWER}}()
    controller.toggle_pause()


# Export functions for LibreOffice
g_exportedScripts = (
    init_{{APPLICATION_NAME_LOWER}},
    start_{{APPLICATION_NAME_LOWER}},
    stop_{{APPLICATION_NAME_LOWER}},
    pause_{{APPLICATION_NAME_LOWER}},
)
"""

# Common patterns for different components
FRAME_UPDATE_PATTERNS = {
    "falling_piece_game": """
            # Move current piece down
            if self.current_piece and self.current_piece.can_move("down"):
                self.current_piece.move("down")
            else:
                self._handle_piece_landed()
                self.current_piece = self._spawn_new_piece()
""",
    "continuous_movement": """
            # Update player/object positions
            self._update_positions()

            # Check collisions
            self._check_collisions()

            # Update score/state
            self._update_game_state()
""",
    "turn_based": """
            # Process AI/computer turns if applicable
            if self._is_computer_turn():
                self._process_computer_turn()
""",
    "simple_animation": """
            # Update animation state
            self._update_animation_frame()
""",
}

RENDER_PATTERNS = {
    "grid_based": """
            self._render_game_board()
            self._render_current_piece()
            self._render_score()
""",
    "cell_coloring": """
            sheet = self.doc.getSheets().getByIndex(0)
            for row in range(self.grid_height):
                for col in range(self.grid_width):
                    cell = sheet.getCellByPosition(col + self.offset_x, row + self.offset_y)
                    if self.grid[row][col] != 0:
                        cell.CellBackColor = self._get_color(self.grid[row][col])
                    else:
                        cell.CellBackColor = -1  # No color
""",
    "minimal": """
            self._draw_screen()
""",
}

END_CONDITIONS = {
    "game_over_flag": "self.game_over",
    "no_more_pieces": "not self.current_piece and not self.can_spawn_piece()",
    "score_threshold": "self.score >= self.win_score",
    "time_limit": "self.elapsed_time >= self.time_limit",
    "manual_only": "False  # Only stop via manual stop call",
}


def generate_timer_loop(
    controller_class_name: str = "GameController",
    state_enum_name: str = "GameState",
    application_name: str = "Game",
    frame_interval: float = 0.5,
    update_pattern: str = "falling_piece_game",
    render_pattern: str = "grid_based",
    end_condition: str = "game_over_flag",
    additional_params: str = "",
    additional_params_desc: str = "",
    init_state_vars: str = "",
    on_start_code: str = "",
    on_stop_code: str = "",
    additional_methods: str = "",
    additional_args: str = "",
) -> str:
    """Generate timer-based loop code from template.

    Args:
        controller_class_name: Name for controller class
        state_enum_name: Name for state enum
        application_name: Application name
        frame_interval: Time between frames in seconds
        update_pattern: Update logic pattern key
        render_pattern: Render logic pattern key
        end_condition: End condition pattern key
        additional_params: Additional __init__ parameters
        additional_params_desc: Description of additional parameters
        init_state_vars: Additional state variable initialization
        on_start_code: Code to run on start
        on_stop_code: Code to run on stop
        additional_methods: Additional class methods
        additional_args: Additional arguments for button handlers

    Returns:
        Generated Python code string
    """
    code = TIMER_LOOP_TEMPLATE

    # Replace placeholders
    code = code.replace("{{CONTROLLER_CLASS_NAME}}", controller_class_name)
    code = code.replace("{{STATE_ENUM_NAME}}", state_enum_name)
    code = code.replace("{{APPLICATION_NAME}}", application_name)
    code = code.replace("{{APPLICATION_NAME_LOWER}}", application_name.lower())
    code = code.replace("{{FRAME_INTERVAL}}", str(frame_interval))
    code = code.replace("{{ADDITIONAL_PARAMS}}", additional_params)
    code = code.replace("{{ADDITIONAL_PARAMS_DESC}}", additional_params_desc)
    code = code.replace("{{INIT_STATE_VARS}}", init_state_vars)
    code = code.replace("{{ON_START_CODE}}", on_start_code)
    code = code.replace("{{ON_STOP_CODE}}", on_stop_code)
    code = code.replace("{{ADDITIONAL_METHODS}}", additional_methods)
    code = code.replace("{{ADDITIONAL_ARGS}}", additional_args)

    # Insert update logic
    update_logic = FRAME_UPDATE_PATTERNS.get(update_pattern, "pass  # Update logic here")
    code = code.replace("{{UPDATE_LOGIC}}", update_logic)

    # Insert render logic
    render_logic = RENDER_PATTERNS.get(render_pattern, "pass  # Render logic here")
    code = code.replace("{{RENDER_LOGIC}}", render_logic)

    # Insert end condition
    end_cond = END_CONDITIONS.get(end_condition, "False")
    code = code.replace("{{END_CONDITION}}", end_cond)

    return code
