"""Workbook inspection API for parse inventory."""

from pathlib import Path
from typing import Any, Literal

from xlsliberator.artifact_inventory import inventory_ods, populate_canonical_inventory
from xlsliberator.extract_excel import extract_workbook
from xlsliberator.extract_vba import extract_vba_modules
from xlsliberator.formula_engine import FormulaDialect
from xlsliberator.formula_semantics import build_formula_ir
from xlsliberator.ir_models import CellType, WorkbookIR
from xlsliberator.validation_models import (
    FormulaIR,
    SourceRef,
    UnsupportedArtifactIR,
    ValidationSeverity,
    WorkbookArtifactIR,
)


def inspect_workbook(
    input_path: str | Path,
    *,
    role: Literal["source", "target"] = "source",
) -> WorkbookArtifactIR:
    """Inspect source workbook artifacts and unsupported coverage."""
    path = Path(input_path)
    if path.suffix.lower() == ".ods":
        return inventory_ods(path, role=role)
    workbook, stats = extract_workbook(path)
    inventory = WorkbookArtifactIR(
        inventory_role=role,
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

    if suffix == ".xlsb":
        inventory.unsupported_artifacts.append(
            UnsupportedArtifactIR(
                source_ref=SourceRef(
                    source_file=str(path),
                    artifact_type="xlsb_semantic_coverage",
                    artifact_id="xlsb-semantic-parsing-incomplete",
                ),
                reason=(
                    "XLSB semantic parsing is incomplete; raw ZIP parts are inventoried, "
                    "but full binary-record coverage is not claimed"
                ),
                severity=ValidationSeverity.ERROR,
                details={
                    "total_formulas": stats.total_formulas,
                    "formulas_extracted": stats.formulas_extracted,
                    "raw_package_inventory": True,
                    "semantic_coverage_complete": False,
                },
            )
        )

    vba_modules: list[dict[str, Any]] = []
    if suffix in {".xlsm", ".xlsb", ".xls"}:
        try:
            modules = extract_vba_modules(path)
            vba_modules = [
                {
                    "name": module.name,
                    "module_type": module.module_type.value,
                    "procedures": module.procedures,
                }
                for module in modules
            ]
            inventory.metadata["vba_modules"] = vba_modules
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

    return populate_canonical_inventory(inventory, path, vba_modules=vba_modules)


def inventory_to_dict(inventory: WorkbookArtifactIR) -> dict[str, Any]:
    """Return a JSON-serializable inventory dict."""
    return inventory.model_dump(mode="json")


def _collect_formula_inventory(workbook: WorkbookIR) -> list[FormulaIR]:
    formulas: list[FormulaIR] = []
    for sheet in workbook.sheets:
        for cell in sheet.cells:
            if cell.cell_type != CellType.FORMULA or not cell.formula:
                continue
            source_ref = SourceRef(
                source_file=workbook.file_path,
                sheet=sheet.name,
                cell_range=cell.address,
                artifact_type="formula",
                artifact_id=f"formula:{sheet.name}!{cell.address}",
            )
            formulas.append(
                build_formula_ir(
                    source_ref=source_ref,
                    formula=cell.formula,
                    dialect=FormulaDialect.EXCEL_A1,
                    array_metadata=cell.formula_metadata,
                    calculation_settings=dict(workbook.metadata.get("calculation_settings") or {}),
                    calculation_order={
                        "source_inventory_index": len(formulas),
                        "declared_chain_position": None,
                        "status": "runtime_evidence_required",
                    },
                )
            )
    for named_range in workbook.named_ranges:
        if not named_range.refers_to.startswith("="):
            continue
        scope = named_range.scope or "workbook"
        source_ref = SourceRef(
            source_file=workbook.file_path,
            sheet=named_range.scope,
            artifact_type="defined_name_formula",
            artifact_id=f"formula:name:{scope}:{named_range.name}",
        )
        formulas.append(
            build_formula_ir(
                source_ref=source_ref,
                formula=named_range.refers_to,
                dialect=FormulaDialect.EXCEL_A1,
                name_context=named_range.name,
                calculation_settings=dict(workbook.metadata.get("calculation_settings") or {}),
                calculation_order={
                    "source_inventory_index": len(formulas),
                    "declared_chain_position": None,
                    "status": "not_applicable_to_defined_name",
                },
            )
        )
    return formulas
