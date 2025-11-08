"""API module for Excel to ODS conversion (Phase 6.2 - Hybrid Approach).

Strategic Decision (2025-11-07): Use LibreOffice native conversion + VBA translation.
Architecture: Excel → soffice native → ODS + VBA extraction → LLM translation → Embed macros
"""

import os
import time
from pathlib import Path

from loguru import logger

from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.extract_excel import extract_workbook
from xlsliberator.extract_vba import extract_vba_modules
from xlsliberator.fix_native_ods import post_process_native_ods
from xlsliberator.report import ConversionReport
from xlsliberator.vba2py_uno import translate_vba_to_python


class ConversionError(Exception):
    """Raised when conversion fails."""


def convert_native(
    input_path: Path,
    output_path: Path,
) -> None:
    """Convert Excel to ODS using LibreOffice native conversion via UNO.

    Args:
        input_path: Path to input Excel file
        output_path: Path for output ODS file

    Raises:
        ConversionError: If native conversion fails

    Note:
        Uses UNO bridge for conversion to avoid conflicts with persistent UNO server.
        This ensures the same LibreOffice instance can be used for formula repair.
    """
    from xlsliberator.uno_conn import UnoCtx

    logger.info(f"Running LibreOffice native conversion via UNO: {input_path.name}")

    try:
        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Remove output file if it exists (UNO Overwrite flag doesn't always work)
        if output_path.exists():
            output_path.unlink()

        # Convert using UNO (avoids subprocess conflicts)
        with UnoCtx() as ctx:
            # Load Excel file
            input_url = f"file://{input_path.absolute()}"

            # LoadComponentFromURL with import filter
            doc = ctx.desktop.loadComponentFromURL(input_url, "_blank", 0, ())

            # Store as ODS
            output_url = f"file://{output_path.absolute()}"

            # Store filter for ODS format
            from com.sun.star.beans import PropertyValue

            store_props = (PropertyValue(Name="FilterName", Value="calc8"),)  # ODS format

            doc.storeToURL(output_url, store_props)
            doc.close(True)

        if not output_path.exists():
            raise ConversionError(
                f"Native conversion succeeded but output not found: {output_path}"
            )

        logger.success(f"Native conversion complete: {output_path}")

    except Exception as e:
        raise ConversionError(f"Native conversion error: {e}") from e


def convert(
    input_path: str | Path,
    output_path: str | Path,
    *,
    locale: str = "en-US",
    strict: bool = False,
    embed_macros: bool = True,
) -> ConversionReport:
    """Convert Excel file to LibreOffice Calc ODS format (Hybrid Approach).

    Args:
        input_path: Path to input Excel file (.xlsx, .xlsm, .xlsb, .xls)
        output_path: Path for output ODS file
        locale: Target locale for formulas (note: native conversion handles this)
        strict: If True, fail on any errors; if False, continue with warnings
        embed_macros: If True, translate and embed VBA macros

    Returns:
        ConversionReport with conversion statistics and results

    Raises:
        ConversionError: If conversion fails (in strict mode)

    Note:
        Phase 6.2 implementation - Hybrid approach:
        1. LibreOffice native conversion (100% formula equivalence)
        2. VBA extraction from original Excel
        3. VBA→Python-UNO translation with LLM
        4. Embed Python macros into native ODS
    """
    input_path = Path(input_path)
    output_path = Path(output_path)

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
        convert_native(input_path, output_path)

        # Step 1.5: Post-process native ODS to fix known bugs
        logger.info("Step 1.5: Post-processing native ODS (fix formulas & ranges)...")
        post_stats = post_process_native_ods(input_path, output_path)
        report.formulas_fixed = post_stats.get("formulas_fixed", 0)

        # Extract metadata for reporting (from original Excel)
        logger.info("Extracting metadata for report...")
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

        # Step 2: Extract VBA from original Excel (if embed_macros=True)
        vba_modules = []
        if embed_macros and input_path.suffix.lower() in [".xlsm", ".xlsb", ".xls"]:
            logger.info("Step 2: Extracting VBA macros...")
            try:
                vba_modules = extract_vba_modules(input_path)
                report.vba_modules = len(vba_modules)
                report.vba_procedures = sum(len(m.procedures) for m in vba_modules)

                if vba_modules:
                    logger.info(f"Found {len(vba_modules)} VBA modules")
                else:
                    logger.info("No VBA macros found")
            except Exception as e:
                msg = f"VBA extraction failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

        # Step 3: Translate VBA to Python-UNO (if VBA found)
        python_modules = {}
        if vba_modules:
            logger.info("Step 3: Translating VBA to Python-UNO with LLM...")

            # Check if LLM is available
            use_llm = bool(os.environ.get("ANTHROPIC_API_KEY"))
            if not use_llm:
                logger.warning(
                    "No ANTHROPIC_API_KEY set - VBA translation will use rule-based fallback"
                )

            try:
                for vba_module in vba_modules:
                    # Translate the entire module source code
                    module_name = f"{vba_module.name}.py"
                    result = translate_vba_to_python(vba_module.source_code, use_llm=use_llm)

                    if result.python_code:
                        python_modules[module_name] = result.python_code
                        logger.debug(f"Translated: {module_name}")

                    # Collect warnings
                    for warning in result.warnings:
                        report.warnings.append(f"VBA translation warning: {warning}")

                logger.success(f"Translated {len(python_modules)} Python modules")

            except Exception as e:
                msg = f"VBA translation failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

        # Step 4: Embed Python macros into native ODS
        if python_modules:
            logger.info("Step 4: Embedding Python macros into ODS...")
            try:
                embed_python_macros(output_path, python_modules)
                logger.success(f"Embedded {len(python_modules)} Python modules")
            except Exception as e:
                msg = f"Macro embedding failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

        # Step 5: Test formula equivalence
        logger.info("Step 5: Testing formula equivalence...")
        try:
            from xlsliberator.testing_lo import compare_excel_calc

            test_result = compare_excel_calc(input_path, output_path)
            report.formulas_matching = test_result.matching
            report.formulas_mismatching = test_result.mismatching
            report.formula_match_rate = test_result.match_rate

            logger.success(
                f"Formula equivalence: {test_result.matching}/{test_result.formula_cells} "
                f"({test_result.match_rate:.2f}%)"
            )

            # Log first few mismatches as warnings
            for mismatch in test_result.mismatches[:5]:
                msg = (
                    f"Formula mismatch at {mismatch['sheet']}!{mismatch['cell']}: "
                    f"Excel={mismatch['excel_value']} vs Calc={mismatch['calc_value']}"
                )
                report.warnings.append(msg)

        except Exception as e:
            msg = f"Formula equivalence testing failed: {e}"
            logger.warning(msg)
            report.warnings.append(msg)

        # Conversion successful
        report.success = True
        report.duration_seconds = time.time() - start_time

        logger.success(f"Hybrid conversion completed in {report.duration_seconds:.2f}s")

        return report

    except Exception as e:
        report.success = False
        report.duration_seconds = time.time() - start_time
        error_msg = f"Conversion failed: {e}"
        report.errors.append(error_msg)
        logger.error(error_msg)

        if strict:
            raise ConversionError(error_msg) from e

        return report
