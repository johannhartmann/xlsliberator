"""Docker-only Excel to LibreOffice Calc conversion API."""

import time
import warnings
import zipfile
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from loguru import logger

from xlsliberator.extract_excel import extract_workbook
from xlsliberator.primitives import (
    extract_vba_project,
    native_convert_workbook,
    upsert_python_modules,
)
from xlsliberator.report import ConversionReport


class ConversionError(Exception):
    """Raised when conversion fails."""


NATIVE_CONVERSION_TIMEOUT_SECONDS = 120
ProgressCallback = Callable[[str, str, dict[str, Any]], None]


def _emit_progress(
    callback: ProgressCallback | None,
    phase: str,
    message: str,
    details: dict[str, Any] | None = None,
) -> None:
    """Emit progress without letting callback failures break conversion."""
    if callback is None:
        return
    try:
        callback(phase, message, details or {})
    except Exception as e:
        logger.warning(f"Progress callback failed: {e}")


def convert_native(
    input_path: Path,
    output_path: Path,
    *,
    user_installation_dir: str | Path | None = None,
    uno_port: int | None = None,
) -> None:
    """Convert Excel to ODS in the authoritative LibreOffice Docker runtime.

    Args:
        input_path: Path to input Excel file
        output_path: Path for output ODS file

    Raises:
        ConversionError: If native conversion fails

    ``user_installation_dir`` and ``uno_port`` remain accepted for API compatibility,
    but are intentionally ignored: host office profiles and host UNO sockets are
    outside the supported runtime boundary.
    """
    del user_installation_dir, uno_port
    logger.info(f"Running Docker-only LibreOffice conversion: {input_path.name}")
    result = native_convert_workbook(
        input_path,
        output_path,
        timeout_seconds=NATIVE_CONVERSION_TIMEOUT_SECONDS,
    )
    if not result.success:
        detail = "; ".join(result.errors) or result.status.value
        raise ConversionError(f"LibreOffice Docker runtime conversion failed: {detail}")
    image_id = str(result.runtime_identity.get("image_id") or "unknown")
    logger.success(f"Docker LibreOffice conversion complete: {output_path} ({image_id})")


def convert(
    input_path: str | Path,
    output_path: str | Path,
    *,
    locale: str = "en-US",
    strict: bool = False,
    embed_macros: bool = False,
    use_agent: bool = False,
    python_modules: Mapping[str, str] | None = None,
    validate_macro_execution: bool = False,
    allow_global_macro_security_change: bool = False,
    progress_callback: ProgressCallback | None = None,
    user_installation_dir: str | Path | None = None,
) -> ConversionReport:
    """Convert an Excel file to LibreOffice Calc ODS format.

    Args:
        input_path: Path to input Excel file (.xlsx, .xlsm, .xlsb, .xls)
        output_path: Path for output ODS file
        locale: Target locale for formulas (note: native conversion handles this)
        strict: If True, fail on any errors; if False, continue with warnings
        embed_macros: If True, require and embed supplied target-native Python modules
        use_agent: Deprecated compatibility flag; model orchestration is external
        python_modules: Agent-produced target-native Python/UNO modules to embed
        validate_macro_execution: If True, run macro execution validation when safe
        allow_global_macro_security_change: Explicit opt-in for legacy global macro security change
        progress_callback: Optional callback for ordered conversion progress events
        user_installation_dir: Optional isolated LibreOffice profile directory

    Returns:
        ConversionReport with conversion statistics and results

    Raises:
        ConversionError: If conversion fails (in strict mode)

    LibreOffice and PyUNO execution is confined to the pinned Docker runtime.
    Runtime validation and differential equivalence are separate certification
    gates and are never inferred from conversion success.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)
    if use_agent:
        warnings.warn(
            "Embedded model orchestration was removed; use xlsliberator-swe and pass "
            "python_modules to the deterministic converter",
            DeprecationWarning,
            stacklevel=2,
        )

    start_time = time.time()

    # Initialize report
    report = ConversionReport(
        input_file=str(input_path),
        output_file=str(output_path),
        success=False,
        locale=locale,
    )

    logger.info(f"Starting hybrid conversion: {input_path} → {output_path}")

    try:
        # Step 1: LibreOffice Native Conversion (formulas, data, formatting)
        logger.info("Step 1: LibreOffice native conversion...")
        _emit_progress(progress_callback, "converting", "Running LibreOffice native conversion")
        if user_installation_dir is None:
            convert_native(input_path, output_path)
        else:
            convert_native(
                input_path,
                output_path,
                user_installation_dir=user_installation_dir,
            )

        # Pure package post-processing; this does not import or start UNO/LibreOffice.
        if input_path.suffix.lower() != ".xls":
            logger.info("Step 1.5: Post-processing native ODS formulas and ranges...")
            _emit_progress(progress_callback, "repairing", "Repairing formulas and named ranges")
            try:
                from xlsliberator.fix_native_ods import post_process_native_ods

                post_stats = post_process_native_ods(input_path, output_path)
                report.formulas_fixed = post_stats.get("formulas_fixed", 0)
            except Exception as exc:
                msg = f"Post-processing failed: {exc}"
                logger.warning(msg)
                report.warnings.append(msg)
        else:
            logger.info("Skipping package post-processing for legacy .xls input")

        # Extract metadata for reporting (from original Excel)
        logger.info("Extracting metadata for report...")
        _emit_progress(progress_callback, "analyzing", "Extracting workbook metadata")
        try:
            wb_ir, _ = extract_workbook(input_path)
            report.total_cells = wb_ir.total_cells
            report.total_formulas = wb_ir.total_formulas
            report.named_ranges = len(wb_ir.named_ranges)
            report.sheet_count = wb_ir.sheet_count
            report.formulas_translated = wb_ir.total_formulas  # Native conversion handles this

            logger.success(
                f"Native conversion: {report.total_cells:,} cells, "
                f"{report.total_formulas:,} formulas, "
                f"{report.sheet_count} sheets"
            )
        except Exception as e:
            msg = f"Metadata extraction failed: {e}"
            logger.warning(msg)
            report.warnings.append(msg)
            # Continue without metadata
            logger.info("Continuing without metadata extraction")

        # Step 2: Extract VBA from the original source for explicit loss accounting.
        vba_modules = []
        if input_path.suffix.lower() in [".xlsm", ".xlsb", ".xls"]:
            logger.info("Step 2: Extracting VBA macros...")
            _emit_progress(progress_callback, "extracting_vba", "Extracting VBA macros")
            extraction = extract_vba_project(input_path)
            if extraction.success:
                vba_modules = extraction.modules
                report.vba_modules = len(vba_modules)
                report.vba_procedures = sum(len(m.procedures) for m in vba_modules)

                if vba_modules:
                    logger.info(f"Found {len(vba_modules)} VBA modules")
                else:
                    logger.info("No VBA macros found")
            else:
                msg = f"VBA extraction failed: {'; '.join(extraction.errors)}"
                logger.warning(msg)
                report.errors.append(msg)

        supplied_modules = dict(python_modules or {})
        if vba_modules and not supplied_modules:
            message = (
                "Source VBA was extracted but not migrated; model orchestration belongs to "
                "xlsliberator-swe"
            )
            if embed_macros:
                report.errors.append(message)
            else:
                report.warnings.append(message)

        # Step 3: Embed caller-supplied target-native Python modules.
        if supplied_modules:
            logger.info("Step 4: Embedding Python macros into ODS...")
            _emit_progress(progress_callback, "embedding", "Embedding supplied Python modules")
            upsert = upsert_python_modules(output_path, supplied_modules)
            if upsert.success:
                logger.success(f"Embedded {len(supplied_modules)} Python modules")
            else:
                report.errors.extend(upsert.errors)

        # Step 4.5: Global macro security changes are disabled by default.
        if allow_global_macro_security_change:
            msg = "Global macro security mutation is unsupported in the Docker-only runtime"
            logger.warning(msg)
            report.warnings.append(msg)
        elif supplied_modules:
            msg = (
                "Skipped global macro security change; runtime macro execution validation "
                "requires an isolated office profile"
            )
            logger.info(msg)
            report.warnings.append(msg)

        if supplied_modules:
            # Step 4.6: Validate embedded Python macros
            logger.info("Step 4.6: Validating embedded Python macros...")
            _emit_progress(
                progress_callback, "verifying_macros", "Validating embedded Python macros"
            )
            try:
                from xlsliberator.python_macro_manager import validate_all_embedded_macros

                validation_summary = validate_all_embedded_macros(output_path)
                report.macros_validated = validation_summary.total_modules
                report.macros_syntax_valid = validation_summary.valid_syntax
                report.macros_syntax_errors = validation_summary.syntax_errors
                report.macros_with_exported_scripts = validation_summary.has_exported_scripts
                report.macros_missing_exported_scripts = validation_summary.missing_exported_scripts

                logger.success(
                    f"Macro validation: {validation_summary.valid_syntax}/"
                    f"{validation_summary.total_modules} valid"
                )

                # Log validation warnings
                for module_name, val_result in validation_summary.validation_details.items():
                    if not val_result.valid:
                        for error in val_result.errors:
                            msg = f"Macro validation error in {module_name}: {error}"
                            report.errors.append(msg)
                    for warning in val_result.warnings:
                        msg = f"Macro validation warning in {module_name}: {warning}"
                        report.warnings.append(msg)

            except Exception as e:
                msg = f"Macro validation failed: {e}"
                logger.warning(msg)
                report.errors.append(msg)

            # Step 4.7: Test macro execution without embedded model repair.
            if not validate_macro_execution and not allow_global_macro_security_change:
                msg = (
                    "Macro execution validation skipped because no isolated runtime profile "
                    "is wired into the conversion pipeline"
                )
                logger.info(msg)
                report.warnings.append(msg)
            else:
                logger.info("Step 4.7: Testing macro execution...")
                _emit_progress(progress_callback, "verifying_macros", "Testing macro execution")
                try:
                    from xlsliberator.python_macro_manager import test_all_macros_safe

                    execution_summary = test_all_macros_safe(output_path)
                    report.macro_functions_tested = execution_summary.total_functions
                    report.macro_functions_passed = execution_summary.successful
                    report.macro_functions_failed = execution_summary.failed
                    report.macro_functions_skipped = execution_summary.skipped

                    logger.success(
                        f"Macro execution: {execution_summary.successful}/"
                        f"{execution_summary.total_functions} passed"
                    )

                    for uri, exec_result in execution_summary.execution_details.items():
                        if not exec_result.success:
                            msg = f"Macro execution failed: {uri}: {exec_result.error}"
                            report.errors.append(msg)

                except Exception as e:
                    msg = f"Macro execution testing failed: {e}"
                    logger.warning(msg)
                    report.errors.append(msg)

            # Step 4.8: Agent-based GUI validation
            logger.info("Step 4.8: Running agent-based GUI validation...")
            _emit_progress(progress_callback, "verifying_gui", "Verifying GUI and event bindings")
            try:
                from xlsliberator.agent_validator import validate_document_with_agent_sync

                agent_result = validate_document_with_agent_sync(output_path)
                report.agent_validation_run = True
                report.agent_macros_validated = agent_result.macros_validated
                report.agent_macros_valid = agent_result.macros_valid
                report.agent_functions_found = agent_result.functions_found
                report.agent_buttons_found = agent_result.buttons_found
                report.agent_buttons_with_handlers = agent_result.buttons_with_handlers
                report.agent_cells_readable = agent_result.cells_readable

                if agent_result.success:
                    logger.success(
                        f"Agent validation: {agent_result.macros_valid} macros, "
                        f"{agent_result.functions_found} functions, "
                        f"{agent_result.buttons_with_handlers} button handlers"
                    )
                else:
                    logger.warning(
                        f"Agent validation completed with issues: "
                        f"{len(agent_result.warnings)} warnings, "
                        f"{len(agent_result.errors)} errors"
                    )

                # Add warnings/errors to report
                for warning in agent_result.warnings:
                    report.warnings.append(f"Agent validation: {warning}")
                for error in agent_result.errors:
                    report.errors.append(f"Agent validation: {error}")

            except Exception as e:
                msg = f"Agent validation failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

        if not output_path.is_file():
            report.errors.append("Conversion did not produce an output file")
        elif not zipfile.is_zipfile(output_path):
            report.errors.append("Conversion output is not a valid ODS ZIP package")
        report.success = (
            output_path.is_file() and zipfile.is_zipfile(output_path) and not report.errors
        )
        report.duration_seconds = time.time() - start_time
        if report.success:
            _emit_progress(
                progress_callback,
                "completed",
                "Conversion complete",
                {"output": str(output_path)},
            )
            logger.success(f"Hybrid conversion completed in {report.duration_seconds:.2f}s")
        else:
            message = "; ".join(report.errors)
            _emit_progress(progress_callback, "failed", message)
            logger.error(f"Hybrid conversion failed: {message}")
            if strict:
                raise ConversionError(message)

        return report

    except Exception as e:
        report.success = False
        report.duration_seconds = time.time() - start_time
        error_msg = f"Conversion failed: {e}"
        report.errors.append(error_msg)
        _emit_progress(progress_callback, "failed", error_msg)
        logger.error(error_msg)

        if strict:
            raise ConversionError(error_msg) from e

        return report
