"""Workbook inspection API for parse inventory."""

from pathlib import Path
from typing import Any

from xlsliberator.extract_excel import extract_workbook
from xlsliberator.extract_vba import extract_vba_modules
from xlsliberator.ir_models import CellType, WorkbookIR
from xlsliberator.validation_models import (
    FormulaIR,
    SourceRef,
    UnsupportedArtifactIR,
    ValidationSeverity,
    WorkbookArtifactIR,
)


def inspect_workbook(input_path: str | Path) -> WorkbookArtifactIR:
    """Inspect source workbook artifacts and unsupported coverage."""
    path = Path(input_path)
    workbook, stats = extract_workbook(path)
    inventory = WorkbookArtifactIR(
        workbook=workbook,
        formulas=_collect_formula_inventory(workbook),
        metadata={
            "extraction_stats": stats.model_dump(),
            "source_suffix": path.suffix.lower(),
        },
    )

    suffix = path.suffix.lower()
    if suffix == ".xls":
        inventory.unsupported_artifacts.append(
            UnsupportedArtifactIR(
                source_ref=SourceRef(
                    source_file=str(path),
                    artifact_type="legacy_xls_biff",
                    artifact_id="legacy-xls-incomplete",
                ),
                reason="legacy XLS BIFF parsing incomplete; cells, formulas, controls, and macros are not fully enumerated",
                severity=ValidationSeverity.ERROR,
                details={"file_format": "xls"},
            )
        )

    if suffix == ".xlsb" and stats.formulas_extracted < stats.total_formulas:
        inventory.unsupported_artifacts.append(
            UnsupportedArtifactIR(
                source_ref=SourceRef(
                    source_file=str(path),
                    artifact_type="xlsb_formula",
                    artifact_id="xlsb-limited-formula-support",
                ),
                reason="XLSB formula extraction is limited by the current reader",
                severity=ValidationSeverity.WARNING,
                details={
                    "total_formulas": stats.total_formulas,
                    "formulas_extracted": stats.formulas_extracted,
                },
            )
        )

    if suffix in {".xlsm", ".xlsb", ".xls"}:
        try:
            modules = extract_vba_modules(path)
            inventory.metadata["vba_modules"] = [
                {
                    "name": module.name,
                    "module_type": module.module_type.value,
                    "procedures": module.procedures,
                }
                for module in modules
            ]
        except Exception as exc:
            inventory.unsupported_artifacts.append(
                UnsupportedArtifactIR(
                    source_ref=SourceRef(
                        source_file=str(path),
                        artifact_type="vba_project",
                        artifact_id="vba-extraction-failed",
                    ),
                    reason=f"VBA extraction failed: {exc}",
                    severity=ValidationSeverity.WARNING,
                    details={"exception_type": type(exc).__name__},
                )
            )

    return inventory


def inventory_to_dict(inventory: WorkbookArtifactIR) -> dict[str, Any]:
    """Return a JSON-serializable inventory dict."""
    return inventory.model_dump(mode="json")


def _collect_formula_inventory(workbook: WorkbookIR) -> list[FormulaIR]:
    formulas: list[FormulaIR] = []
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if cell.cell_type != CellType.FORMULA or not cell.formula:
                continue
            formulas.append(
                FormulaIR(
                    source_ref=SourceRef(
                        source_file=workbook.file_path,
                        sheet=sheet.name,
                        cell_range=cell.address,
                        artifact_type="formula",
                        artifact_id=f"{sheet.name}!{cell.address}",
                    ),
                    formula_text=cell.formula,
                    dialect="excel_a1",
                )
            )
    return formulas
