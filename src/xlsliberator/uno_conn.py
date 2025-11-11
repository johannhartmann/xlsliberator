"""LibreOffice UNO connection management and helper functions."""

import socket
import subprocess
import time
from pathlib import Path
from typing import Any

from loguru import logger


class UnoConnectionError(Exception):
    """Raised when UNO connection fails."""


class UnoCtx:
    """LibreOffice UNO connection context manager."""

    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 2002,
        timeout: int = 10,
        manage_libreoffice: bool = True,
        use_gui: bool = False,
    ) -> None:
        """Initialize UNO connection context.

        Args:
            host: LibreOffice host address
            port: LibreOffice UNO port
            timeout: Connection timeout in seconds
            manage_libreoffice: If True, start/stop LibreOffice process automatically
            use_gui: If True, use GUI mode with xvfb (enables XScriptProvider for macros)
        """
        self.host = host
        self.port = port
        self.timeout = timeout
        self.manage_libreoffice = manage_libreoffice
        self.use_gui = use_gui
        self.local_context: Any = None
        self.resolver: Any = None
        self.component_context: Any = None
        self.desktop: Any = None
        self._connected = False
        self._libreoffice_process: subprocess.Popen | None = None
        self._xvfb_process: subprocess.Popen | None = None
        self._display: str | None = None

    def __enter__(self) -> "UnoCtx":
        """Enter context and establish connection."""
        self.connect()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Exit context and close connection."""
        self.disconnect()

    def _start_libreoffice(self) -> None:
        """Start LibreOffice process (headless or GUI mode with xvfb)."""
        import os
        import shutil

        mode = "GUI mode with xvfb" if self.use_gui else "headless"
        logger.info(f"Starting LibreOffice in {mode} on port {self.port}")

        # Kill any existing LibreOffice processes
        try:
            subprocess.run(["pkill", "-9", "soffice"], check=False, capture_output=True)
            time.sleep(1)
        except Exception:
            pass

        # Start Xvfb if GUI mode requested
        if self.use_gui:
            if not shutil.which("Xvfb"):
                logger.warning("Xvfb not found, falling back to headless mode")
                self.use_gui = False
            else:
                # Find available display number
                for display_num in range(99, 200):
                    display = f":{display_num}"
                    # Check if display is already in use
                    lock_file = f"/tmp/.X{display_num}-lock"
                    if not Path(lock_file).exists():
                        self._display = display
                        break
                else:
                    logger.warning("No available X display found, falling back to headless")
                    self.use_gui = False

                if self.use_gui and self._display:
                    # Start Xvfb on the selected display
                    try:
                        self._xvfb_process = subprocess.Popen(
                            [
                                "Xvfb",
                                self._display,
                                "-screen",
                                "0",
                                "1280x1024x24",
                                "-nolisten",
                                "tcp",
                            ],
                            stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL,
                        )
                        time.sleep(1)  # Give Xvfb time to start
                        logger.debug(f"Started Xvfb on display {self._display}")
                    except Exception as e:
                        logger.warning(f"Failed to start Xvfb: {e}, falling back to headless")
                        self.use_gui = False
                        if self._xvfb_process:
                            self._xvfb_process.kill()
                            self._xvfb_process = None

        # Build LibreOffice command
        if self.use_gui and self._display:
            cmd = ["soffice"]
            env = os.environ.copy()
            env["DISPLAY"] = self._display
        else:
            cmd = ["soffice", "--headless"]
            env = None

        # Add common arguments
        cmd.extend(
            [
                f"--accept=socket,host={self.host},port={self.port};urp;",
                "--norestore",
                "--nofirststartwizard",
            ]
        )

        try:
            self._libreoffice_process = subprocess.Popen(
                cmd,
                env=env,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            logger.debug(f"LibreOffice process started (PID: {self._libreoffice_process.pid})")

            # Wait for LibreOffice to start accepting connections
            for _attempt in range(30):  # Try for 30 seconds
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(1)
                result = sock.connect_ex((self.host, self.port))
                sock.close()
                if result == 0:
                    logger.success("LibreOffice is ready")
                    return
                time.sleep(1)

            raise UnoConnectionError("LibreOffice failed to start accepting connections")

        except FileNotFoundError as e:
            raise UnoConnectionError("soffice command not found. Is LibreOffice installed?") from e

    def _stop_libreoffice(self) -> None:
        """Stop LibreOffice and Xvfb processes."""
        if self._libreoffice_process:
            logger.info("Stopping LibreOffice process")
            try:
                self._libreoffice_process.terminate()
                self._libreoffice_process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                logger.warning("LibreOffice did not terminate gracefully, killing")
                self._libreoffice_process.kill()
                self._libreoffice_process.wait()
            except Exception as e:
                logger.warning(f"Error stopping LibreOffice: {e}")
            finally:
                self._libreoffice_process = None

        # Stop Xvfb if we started it
        if self._xvfb_process:
            logger.debug("Stopping Xvfb process")
            try:
                self._xvfb_process.terminate()
                self._xvfb_process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self._xvfb_process.kill()
                self._xvfb_process.wait()
            except Exception as e:
                logger.warning(f"Error stopping Xvfb: {e}")
            finally:
                self._xvfb_process = None
                self._display = None

    def connect(self) -> None:
        """Establish connection to LibreOffice UNO."""
        try:
            # Start LibreOffice if managing it
            if self.manage_libreoffice:
                self._start_libreoffice()

            # Import UNO modules (only when actually connecting)
            import uno
            from com.sun.star.connection import NoConnectException

            logger.info(f"Connecting to LibreOffice at {self.host}:{self.port}")

            # Get local component context
            self.local_context = uno.getComponentContext()

            # Get resolver
            self.resolver = self.local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", self.local_context
            )

            # Check if LibreOffice is running
            if not self.manage_libreoffice:
                try:
                    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                    sock.settimeout(self.timeout)
                    result = sock.connect_ex((self.host, self.port))
                    sock.close()
                    if result != 0:
                        raise UnoConnectionError(
                            f"LibreOffice not reachable at {self.host}:{self.port}. "
                            f"Start with: soffice --headless --accept='socket,host={self.host},"
                            f"port={self.port};urp;'"
                        )
                except OSError as e:
                    raise UnoConnectionError(f"Socket error: {e}") from e

            # Connect to LibreOffice
            uno_url = (
                f"uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext"
            )

            try:
                self.component_context = self.resolver.resolve(uno_url)
            except NoConnectException as e:
                raise UnoConnectionError(
                    f"Failed to connect to LibreOffice UNO at {uno_url}"
                ) from e

            # Get desktop service
            self.desktop = self.component_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.component_context
            )

            self._connected = True
            logger.success("Connected to LibreOffice UNO")

        except ImportError as e:
            if self.manage_libreoffice:
                self._stop_libreoffice()
            raise UnoConnectionError(
                "UNO Python bindings not available. Install LibreOffice SDK."
            ) from e
        except Exception as e:
            if self.manage_libreoffice:
                self._stop_libreoffice()
            raise UnoConnectionError(f"Failed to connect to UNO: {e}") from e

    def disconnect(self) -> None:
        """Close connection to LibreOffice UNO."""
        if self._connected:
            logger.info("Disconnecting from LibreOffice UNO")
            # UNO connections are automatically cleaned up
            self._connected = False
            self.desktop = None
            self.component_context = None
            self.resolver = None
            self.local_context = None
            logger.success("Disconnected from LibreOffice UNO")

        # Stop LibreOffice if we started it
        if self.manage_libreoffice:
            self._stop_libreoffice()

    @property
    def is_connected(self) -> bool:
        """Check if connected to LibreOffice."""
        return self._connected


def connect_lo(host: str = "127.0.0.1", port: int = 2002, timeout: int = 10) -> UnoCtx:
    """Connect to LibreOffice UNO.

    Args:
        host: LibreOffice host address
        port: LibreOffice UNO port
        timeout: Connection timeout in seconds

    Returns:
        UnoCtx context manager

    Raises:
        UnoConnectionError: If connection fails
    """
    ctx = UnoCtx(host=host, port=port, timeout=timeout)
    ctx.connect()
    return ctx


def new_calc(ctx: UnoCtx) -> Any:
    """Create a new Calc spreadsheet document.

    Args:
        ctx: UNO connection context

    Returns:
        Calc document object

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    logger.debug("Creating new Calc document")
    doc = ctx.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
    logger.debug("Created new Calc document")
    return doc


def set_macro_security_level(ctx: UnoCtx, level: int = 0) -> None:
    """Set LibreOffice macro security level.

    Args:
        ctx: UNO connection context
        level: Security level (0=Low, 1=Medium, 2=High, 3=Very High)

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    try:
        # Get configuration provider
        from com.sun.star.beans import PropertyValue

        config_provider = ctx.component_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx.component_context
        )

        # Configuration path for macro security
        config_path = PropertyValue()
        config_path.Name = "nodepath"
        config_path.Value = "/org.openoffice.Office.Common/Security/Scripting"

        # Get update access to the configuration
        config_access = config_provider.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationUpdateAccess", (config_path,)
        )

        # Set MacroSecurityLevel (0=Low, 1=Medium, 2=High, 3=Very High)
        config_access.setPropertyValue("MacroSecurityLevel", level)

        # Commit changes
        config_access.commitChanges()

        level_names = {0: "Low", 1: "Medium", 2: "High", 3: "Very High"}
        logger.success(f"Set macro security level to {level_names.get(level, level)}")

    except Exception as e:
        logger.warning(f"Failed to set macro security level: {e}")
        raise UnoConnectionError(f"Failed to set macro security level: {e}") from e


def open_calc(ctx: UnoCtx, path: str | Path) -> Any:
    """Open an existing Calc spreadsheet document.

    Args:
        ctx: UNO connection context
        path: Path to .ods file

    Returns:
        Calc document object

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    # Convert to file URL
    file_path = Path(path).resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    from com.sun.star.beans import PropertyValue
    from uno import systemPathToFileUrl

    file_url = systemPathToFileUrl(str(file_path))
    logger.debug(f"Opening Calc document: {file_url}")

    # Enable macros automatically (MacroExecutionMode = 4 = ALWAYS_EXECUTE_NO_WARN)
    # This allows Python-UNO macros to run without security warnings
    load_props = (
        PropertyValue(Name="MacroExecutionMode", Value=4),  # ALWAYS_EXECUTE_NO_WARN
    )

    doc = ctx.desktop.loadComponentFromURL(file_url, "_blank", 0, load_props)
    logger.debug("Opened Calc document with macros enabled")
    return doc


def save_as_ods(ctx: UnoCtx, doc: Any, path: str | Path) -> None:
    """Save document as ODS file.

    Args:
        ctx: UNO connection context
        doc: Calc document object
        path: Output path for .ods file

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    # Convert to file URL
    file_path = Path(path).resolve()
    file_path.parent.mkdir(parents=True, exist_ok=True)

    from uno import systemPathToFileUrl

    file_url = systemPathToFileUrl(str(file_path))
    logger.debug(f"Saving document as: {file_url}")

    # Store as ODS format
    from com.sun.star.beans import PropertyValue

    props = (PropertyValue(Name="FilterName", Value="calc8"),)
    doc.storeAsURL(file_url, props)
    logger.debug("Document saved")


def recalc(ctx: UnoCtx, doc: Any) -> None:
    """Recalculate all formulas in document.

    Args:
        ctx: UNO connection context
        doc: Calc document object

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    logger.debug("Recalculating document")
    doc.calculateAll()
    logger.debug("Recalculation complete")


def get_sheet(ctx: UnoCtx, doc: Any, name_or_index: str | int) -> Any:
    """Get sheet by name or index.

    Args:
        ctx: UNO connection context
        doc: Calc document object
        name_or_index: Sheet name (str) or index (int)

    Returns:
        Sheet object

    Raises:
        UnoConnectionError: If not connected
        KeyError: If sheet not found
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    sheets = doc.getSheets()

    if isinstance(name_or_index, str):
        if not sheets.hasByName(name_or_index):
            raise KeyError(f"Sheet not found: {name_or_index}")
        return sheets.getByName(name_or_index)
    else:
        if name_or_index < 0 or name_or_index >= sheets.getCount():
            raise IndexError(f"Sheet index out of range: {name_or_index}")
        return sheets.getByIndex(name_or_index)


def get_cell(ctx: UnoCtx, sheet: Any, address: str) -> Any:
    """Get cell by address (e.g., 'A1').

    Args:
        ctx: UNO connection context
        sheet: Sheet object
        address: Cell address (e.g., 'A1', 'B10')

    Returns:
        Cell object

    Raises:
        UnoConnectionError: If not connected
    """
    if not ctx.is_connected:
        raise UnoConnectionError("Not connected to LibreOffice")

    return sheet.getCellRangeByName(address)
