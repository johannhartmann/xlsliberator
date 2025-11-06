"""Intermediate Representation (IR) models for Excel workbook data."""

from enum import Enum
from typing import Any

from pydantic import BaseModel, Field


class CellType(str, Enum):
    """Cell value type."""

    EMPTY = "empty"
    NUMBER = "number"
    STRING = "string"
    BOOLEAN = "boolean"
    ERROR = "error"
    FORMULA = "formula"


class CellIR(BaseModel):
    """Intermediate representation of a cell."""

    row: int = Field(description="Row index (0-based)")
    col: int = Field(description="Column index (0-based)")
    address: str = Field(description="Cell address (e.g., 'A1')")
    cell_type: CellType = Field(description="Type of cell content")
    value: Any | None = Field(default=None, description="Cell value")
    formula: str | None = Field(default=None, description="Formula string (if applicable)")
    style: dict[str, Any] | None = Field(default=None, description="Cell style/formatting")
    comment: str | None = Field(default=None, description="Cell comment")


class NamedRangeIR(BaseModel):
    """Intermediate representation of a named range."""

    name: str = Field(description="Name of the range")
    scope: str | None = Field(default=None, description="Scope (workbook or sheet name)")
    refers_to: str = Field(description="Range reference formula")
    comment: str | None = Field(default=None, description="Comment/description")


class ChartMetadataIR(BaseModel):
    """Metadata about charts (full parsing in later phase)."""

    chart_id: str = Field(description="Chart identifier")
    chart_type: str | None = Field(default=None, description="Chart type")
    title: str | None = Field(default=None, description="Chart title")
    data_range: str | None = Field(default=None, description="Data source range")


class TableMetadataIR(BaseModel):
    """Metadata about Excel tables (ListObjects)."""

    name: str = Field(description="Table name")
    display_name: str | None = Field(default=None, description="Display name")
    ref: str = Field(description="Table range reference (e.g., 'A1:D10')")
    header_row: bool = Field(default=True, description="Has header row")
    totals_row: bool = Field(default=False, description="Has totals row")
    columns: list[str] = Field(default_factory=list, description="Column names")


class SheetIR(BaseModel):
    """Intermediate representation of a worksheet."""

    name: str = Field(description="Sheet name")
    index: int = Field(description="Sheet index (0-based)")
    visible: bool = Field(default=True, description="Sheet visibility")
    cells: list[CellIR] = Field(default_factory=list, description="All cells")
    tables: list[TableMetadataIR] = Field(default_factory=list, description="Excel tables")
    charts: list[ChartMetadataIR] = Field(default_factory=list, description="Chart metadata")
    max_row: int = Field(default=0, description="Maximum row with data")
    max_col: int = Field(default=0, description="Maximum column with data")

    @property
    def cell_count(self) -> int:
        """Total number of cells."""
        return len(self.cells)

    @property
    def formula_count(self) -> int:
        """Number of cells with formulas."""
        return sum(1 for cell in self.cells if cell.cell_type == CellType.FORMULA)


class WorkbookIR(BaseModel):
    """Intermediate representation of an Excel workbook."""

    file_path: str = Field(description="Source file path")
    file_format: str = Field(description="File format (xlsx, xlsm, xlsb, xls)")
    sheets: list[SheetIR] = Field(default_factory=list, description="All sheets")
    named_ranges: list[NamedRangeIR] = Field(default_factory=list, description="Named ranges")
    has_macros: bool = Field(default=False, description="Has VBA macros (vbaProject.bin)")
    has_external_links: bool = Field(default=False, description="Has external workbook links")
    metadata: dict[str, Any] = Field(default_factory=dict, description="Additional metadata")

    @property
    def sheet_count(self) -> int:
        """Total number of sheets."""
        return len(self.sheets)

    @property
    def total_cells(self) -> int:
        """Total cells across all sheets."""
        return sum(sheet.cell_count for sheet in self.sheets)

    @property
    def total_formulas(self) -> int:
        """Total formulas across all sheets."""
        return sum(sheet.formula_count for sheet in self.sheets)

    def get_sheet_by_name(self, name: str) -> SheetIR | None:
        """Get sheet by name."""
        for sheet in self.sheets:
            if sheet.name == name:
                return sheet
        return None

    def get_sheet_by_index(self, index: int) -> SheetIR | None:
        """Get sheet by index."""
        for sheet in self.sheets:
            if sheet.index == index:
                return sheet
        return None


class ExtractionStats(BaseModel):
    """Statistics about the extraction process."""

    total_cells: int = 0
    total_formulas: int = 0
    formulas_extracted: int = 0
    named_ranges_count: int = 0
    tables_count: int = 0
    charts_count: int = 0
    extraction_time_seconds: float = 0.0
    memory_peak_mb: float = 0.0

    @property
    def formula_extraction_rate(self) -> float:
        """Percentage of formulas successfully extracted."""
        if self.total_formulas == 0:
            return 100.0
        return (self.formulas_extracted / self.total_formulas) * 100.0
