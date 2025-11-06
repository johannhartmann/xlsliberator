"""LibreOffice UNO connection management and helper functions."""

import socket
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
    ) -> None:
        """Initialize UNO connection context.

        Args:
            host: LibreOffice host address
            port: LibreOffice UNO port
            timeout: Connection timeout in seconds
        """
        self.host = host
        self.port = port
        self.timeout = timeout
        self.local_context: Any = None
        self.resolver: Any = None
        self.component_context: Any = None
        self.desktop: Any = None
        self._connected = False

    def __enter__(self) -> "UnoCtx":
        """Enter context and establish connection."""
        self.connect()
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        """Exit context and close connection."""
        self.disconnect()

    def connect(self) -> None:
        """Establish connection to LibreOffice UNO."""
        try:
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
            raise UnoConnectionError(
                "UNO Python bindings not available. Install LibreOffice SDK."
            ) from e
        except Exception as e:
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

    from uno import systemPathToFileUrl

    file_url = systemPathToFileUrl(str(file_path))
    logger.debug(f"Opening Calc document: {file_url}")

    doc = ctx.desktop.loadComponentFromURL(file_url, "_blank", 0, ())
    logger.debug("Opened Calc document")
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
