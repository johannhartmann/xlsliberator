"""XKeyListener template for keyboard input handling."""

KEYBOARD_LISTENER_TEMPLATE = """
# Keyboard Listener Implementation for LibreOffice Python-UNO

import uno
import unohelper
from com.sun.star.awt import XKeyListener
from com.sun.star.awt import Key
from loguru import logger


class {{LISTENER_CLASS_NAME}}(unohelper.Base, XKeyListener):
    \"\"\"Keyboard event listener for {{APPLICATION_NAME}}.

    Handles keyboard input events and updates game/application state accordingly.
    \"\"\"

    def __init__(self, {{CONTROLLER_PARAM}}):
        \"\"\"Initialize keyboard listener.

        Args:
            {{CONTROLLER_PARAM}}: Application controller instance
        \"\"\"
        self.{{CONTROLLER_ATTR}} = {{CONTROLLER_PARAM}}
        self.pressed_keys = set()
        logger.debug("Keyboard listener initialized")

    def keyPressed(self, event):
        \"\"\"Handle key press events.

        Args:
            event: com.sun.star.awt.KeyEvent
        \"\"\"
        key_code = event.KeyCode
        self.pressed_keys.add(key_code)

        try:
            # Key mappings - customize based on application needs
            {{KEY_HANDLERS}}
        except Exception as e:
            logger.error(f"Error handling key press: {e}")

    def keyReleased(self, event):
        \"\"\"Handle key release events.

        Args:
            event: com.sun.star.awt.KeyEvent
        \"\"\"
        self.pressed_keys.discard(event.KeyCode)

    def disposing(self, event):
        \"\"\"Handle listener disposal.

        Args:
            event: Disposal event
        \"\"\"
        logger.debug("Keyboard listener disposed")
        self.pressed_keys.clear()

    def is_key_pressed(self, key_code):
        \"\"\"Check if a specific key is currently pressed.

        Args:
            key_code: Key code to check

        Returns:
            bool: True if key is pressed
        \"\"\"
        return key_code in self.pressed_keys


def register_keyboard_listener(doc, {{CONTROLLER_PARAM}}):
    \"\"\"Register keyboard listener with document.

    Args:
        doc: LibreOffice document instance
        {{CONTROLLER_PARAM}}: Application controller

    Returns:
        Listener instance
    \"\"\"
    listener = {{LISTENER_CLASS_NAME}}({{CONTROLLER_PARAM}})

    try:
        # Get container window from document controller
        controller = doc.getCurrentController()
        frame = controller.getFrame()
        window = frame.getContainerWindow()

        # Add key listener to window
        window.addKeyListener(listener)

        logger.info("Keyboard listener registered successfully")
        return listener

    except Exception as e:
        logger.error(f"Failed to register keyboard listener: {e}")
        raise


# Common key code mappings (LibreOffice Key constants)
KEY_MAPPINGS = {
    # Arrow keys
    "DOWN": Key.DOWN,
    "UP": Key.UP,
    "LEFT": Key.LEFT,
    "RIGHT": Key.RIGHT,

    # Special keys
    "SPACE": Key.SPACE,
    "RETURN": Key.RETURN,
    "ESCAPE": Key.ESCAPE,
    "TAB": Key.TAB,

    # Letter keys (A-Z)
    "A": 65, "B": 66, "C": 67, "D": 68, "E": 69,
    "F": 70, "G": 71, "H": 72, "I": 73, "J": 74,
    "K": 75, "L": 76, "M": 77, "N": 78, "O": 79,
    "P": 80, "Q": 81, "R": 82, "S": 83, "T": 84,
    "U": 85, "V": 86, "W": 87, "X": 88, "Y": 89, "Z": 90,

    # Number keys
    "0": 48, "1": 49, "2": 50, "3": 51, "4": 52,
    "5": 53, "6": 54, "7": 55, "8": 56, "9": 57,

    # Control keys
    "CTRL": Key.MOD1,  # Command on Mac
    "SHIFT": Key.SHIFT,
    "ALT": Key.MOD2,
}
"""

# Example key handler patterns for different use cases
KEY_HANDLER_PATTERNS = {
    "game_movement": """
            if key_code == Key.DOWN:
                self.{{CONTROLLER_ATTR}}.move_down()
            elif key_code == Key.UP:
                self.{{CONTROLLER_ATTR}}.move_up()
            elif key_code == Key.LEFT:
                self.{{CONTROLLER_ATTR}}.move_left()
            elif key_code == Key.RIGHT:
                self.{{CONTROLLER_ATTR}}.move_right()
            elif key_code == Key.SPACE:
                self.{{CONTROLLER_ATTR}}.action()
""",
    "game_with_rotation": """
            if key_code == Key.DOWN:
                if self.{{CONTROLLER_ATTR}}.current_piece:
                    self.{{CONTROLLER_ATTR}}.current_piece.move("down")
            elif key_code == Key.UP:
                if self.{{CONTROLLER_ATTR}}.current_piece:
                    self.{{CONTROLLER_ATTR}}.current_piece.rotate("right")
            elif key_code == Key.LEFT:
                if self.{{CONTROLLER_ATTR}}.current_piece:
                    self.{{CONTROLLER_ATTR}}.current_piece.move("left")
            elif key_code == Key.RIGHT:
                if self.{{CONTROLLER_ATTR}}.current_piece:
                    self.{{CONTROLLER_ATTR}}.current_piece.move("right")
            elif key_code == Key.SPACE:
                if self.{{CONTROLLER_ATTR}}.current_piece:
                    self.{{CONTROLLER_ATTR}}.current_piece.move("drop")
""",
    "wasd_movement": """
            if key_code == 87:  # W
                self.{{CONTROLLER_ATTR}}.move_up()
            elif key_code == 83:  # S
                self.{{CONTROLLER_ATTR}}.move_down()
            elif key_code == 65:  # A
                self.{{CONTROLLER_ATTR}}.move_left()
            elif key_code == 68:  # D
                self.{{CONTROLLER_ATTR}}.move_right()
""",
    "general_control": """
            # Handle key based on application logic
            self.{{CONTROLLER_ATTR}}.handle_key_event(key_code, event)
""",
}


def generate_keyboard_listener(
    listener_class_name: str = "GameKeyListener",
    controller_param: str = "game_controller",
    controller_attr: str = "game",
    key_handler_pattern: str = "game_with_rotation",
) -> str:
    """Generate keyboard listener code from template.

    Args:
        listener_class_name: Name for the listener class
        controller_param: Parameter name for controller
        controller_attr: Attribute name for storing controller
        key_handler_pattern: Key handler pattern name

    Returns:
        Generated Python code string
    """
    code = KEYBOARD_LISTENER_TEMPLATE

    # Replace placeholders
    code = code.replace("{{LISTENER_CLASS_NAME}}", listener_class_name)
    code = code.replace("{{APPLICATION_NAME}}", "Application")
    code = code.replace("{{CONTROLLER_PARAM}}", controller_param)
    code = code.replace("{{CONTROLLER_ATTR}}", controller_attr)

    # Insert key handlers
    handler_code = KEY_HANDLER_PATTERNS.get(
        key_handler_pattern, KEY_HANDLER_PATTERNS["general_control"]
    )
    handler_code = handler_code.replace("{{CONTROLLER_ATTR}}", controller_attr)
    code = code.replace("{{KEY_HANDLERS}}", handler_code)

    return code
