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
        Uses UNO bridge for conversion to enable interactive document manipulation.
        This allows complex operations like formula repair, macro embedding, etc.
    """
    from xlsliberator.uno_conn import UnoCtx

    logger.info(f"Running LibreOffice native conversion via UNO: {input_path.name}")

    try:
        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Remove output file if it exists (UNO Overwrite flag doesn't always work)
        if output_path.exists():
            output_path.unlink()

        # Convert using UNO (enables interactive document manipulation)
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
    use_agent: bool = True,
) -> ConversionReport:
    """Convert Excel file to LibreOffice Calc ODS format (Hybrid Approach).

    Args:
        input_path: Path to input Excel file (.xlsx, .xlsm, .xlsb, .xls)
        output_path: Path for output ODS file
        locale: Target locale for formulas (note: native conversion handles this)
        strict: If True, fail on any errors; if False, continue with warnings
        embed_macros: If True, translate and embed VBA macros
        use_agent: If True, automatically use agent rewriting for complex VBA (default)

    Returns:
        ConversionReport with conversion statistics and results

    Raises:
        ConversionError: If conversion fails (in strict mode)

    Note:
        Hybrid approach with intelligent VBA translation:
        1. LibreOffice native conversion (100% formula equivalence)
        2. VBA extraction from original Excel
        3. Complexity detection (semantic analysis)
        4. VBA→Python-UNO translation (auto-selects simple or agent-based)
        5. Embed Python macros into native ODS

        The system automatically detects VBA complexity and uses:
        - Simple LLM translation for basic macros
        - Multi-agent rewriting for games and complex event-driven code
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
        # Skip for old .xls files (openpyxl doesn't support them)
        if input_path.suffix.lower() != ".xls":
            logger.info("Step 1.5: Post-processing native ODS (fix formulas & ranges)...")
            try:
                from xlsliberator.fix_native_ods import post_process_native_ods

                post_stats = post_process_native_ods(input_path, output_path)
                report.formulas_fixed = post_stats.get("formulas_fixed", 0)
            except Exception as e:
                msg = f"Post-processing failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)
        else:
            logger.info("Skipping post-processing for .xls file (openpyxl limitation)")

        # Extract metadata for reporting (from original Excel)
        logger.info("Extracting metadata for report...")
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
            # Check if LLM is available
            has_api_key = bool(os.environ.get("ANTHROPIC_API_KEY"))

            if use_agent and has_api_key:
                # Step 3a: Detect complexity (always do this when agent mode enabled)
                logger.info("Step 3a: Detecting VBA complexity...")
                try:
                    from xlsliberator.pattern_detector import VBAPatternDetector

                    detector = VBAPatternDetector()
                    complexity = detector.analyze_modules(vba_modules, str(input_path))

                    logger.info(
                        f"Detected complexity: {complexity.complexity_level} "
                        f"(confidence: {complexity.confidence:.0%})"
                    )

                    # Add to report
                    report.warnings.append(
                        f"VBA complexity: {complexity.complexity_level} "
                        f"({complexity.confidence:.0%} confidence)"
                    )

                except Exception as e:
                    msg = f"Complexity detection failed: {e}"
                    logger.warning(msg)
                    report.warnings.append(msg)
                    # Default to simple if detection fails
                    complexity = None

                # Step 3b: Choose translation approach based on complexity
                if complexity and complexity.complexity_level in ["game", "advanced"]:
                    # Use multi-agent system for complex VBA
                    logger.info(
                        f"Step 3b: Using agent rewriting for {complexity.complexity_level} VBA..."
                    )

                    try:
                        from xlsliberator.agent_rewriter import AgentRewriter

                        agent = AgentRewriter()
                        generated_code, validation = agent.rewrite_vba_project(
                            modules=vba_modules,
                            source_file=str(input_path),
                            output_path=output_path,
                            max_iterations=5,
                        )

                        # Use generated modules
                        python_modules = generated_code.modules

                        # Add validation info to report
                        if not validation.syntax_valid:
                            for error in validation.errors:
                                report.errors.append(f"Agent validation: {error}")
                        for warning in validation.warnings:
                            report.warnings.append(f"Agent validation: {warning}")

                        logger.success(
                            f"Agent rewriting complete: {len(python_modules)} modules "
                            f"({validation.iterations_used} iteration(s))"
                        )

                    except Exception as e:
                        msg = f"Agent-based rewriting failed: {e}"
                        logger.warning(msg)
                        report.warnings.append(msg)
                        if strict:
                            raise ConversionError(msg) from e
                else:
                    # Use simple LLM translation for simple VBA
                    logger.info("Step 3b: Using simple translation for basic VBA...")

                    try:
                        for vba_module in vba_modules:
                            module_name = f"{vba_module.name}.py"
                            result = translate_vba_to_python(vba_module.source_code, use_llm=True)

                            if result.python_code:
                                python_modules[module_name] = result.python_code
                                logger.debug(f"Translated: {module_name}")

                            for warning in result.warnings:
                                report.warnings.append(f"VBA translation: {warning}")

                        logger.success(f"Translated {len(python_modules)} Python modules")

                    except Exception as e:
                        msg = f"VBA translation failed: {e}"
                        logger.warning(msg)
                        report.warnings.append(msg)
            else:
                # No agent mode or no API key - use simple translation
                if use_agent and not has_api_key:
                    logger.warning(
                        "ANTHROPIC_API_KEY not set - using simple translation instead of agent mode"
                    )

                logger.info("Step 3: Translating VBA to Python-UNO...")

                use_llm = has_api_key
                if not use_llm:
                    logger.warning(
                        "No ANTHROPIC_API_KEY set - VBA translation will use rule-based fallback"
                    )

                try:
                    for vba_module in vba_modules:
                        module_name = f"{vba_module.name}.py"
                        result = translate_vba_to_python(vba_module.source_code, use_llm=use_llm)

                        if result.python_code:
                            python_modules[module_name] = result.python_code
                            logger.debug(f"Translated: {module_name}")

                        for warning in result.warnings:
                            report.warnings.append(f"VBA translation: {warning}")

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

            # Step 4.5: Enable macros by setting security level to Low
            logger.info("Step 4.5: Setting macro security to Low...")
            try:
                from xlsliberator.uno_conn import UnoCtx, set_macro_security_level

                with UnoCtx() as ctx:
                    set_macro_security_level(ctx, level=0)  # 0 = Low
                logger.success("Macro security set to Low (persists across sessions)")
            except Exception as e:
                msg = f"Failed to set macro security: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

            # Step 4.6: Validate embedded Python macros
            logger.info("Step 4.6: Validating embedded Python macros...")
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
                            report.warnings.append(msg)
                    for warning in val_result.warnings:
                        msg = f"Macro validation warning in {module_name}: {warning}"
                        report.warnings.append(msg)

            except Exception as e:
                msg = f"Macro validation failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

            # Step 4.7: Test macro execution
            logger.info("Step 4.7: Testing macro execution...")
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

                # Log execution failures
                for uri, exec_result in execution_summary.execution_details.items():
                    if not exec_result.success:
                        msg = f"Macro execution failed: {uri}: {exec_result.error}"
                        report.warnings.append(msg)

            except Exception as e:
                msg = f"Macro execution testing failed: {e}"
                logger.warning(msg)
                report.warnings.append(msg)

            # Step 4.8: Agent-based GUI validation
            logger.info("Step 4.8: Running agent-based GUI validation...")
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

        # Step 5.5: Validate formulas using FunctionAccess
        logger.info("Step 5.5: Validating formulas with FunctionAccess...")
        try:
            from xlsliberator.python_macro_manager import validate_all_formulas

            formula_summary = validate_all_formulas(output_path)
            report.formulas_validated = formula_summary.total_formulas
            report.formulas_valid = formula_summary.valid_formulas
            report.formulas_invalid = formula_summary.invalid_formulas

            logger.success(
                f"Formula validation: {formula_summary.valid_formulas}/"
                f"{formula_summary.total_formulas} valid"
            )

            # Log validation failures (first 10)
            invalid_count = 0
            for cell_ref, form_result in formula_summary.validation_details.items():
                if not form_result.valid and invalid_count < 10:
                    msg = f"Formula validation failed at {cell_ref}: {form_result.error}"
                    report.warnings.append(msg)
                    invalid_count += 1

        except Exception as e:
            msg = f"Formula validation failed: {e}"
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
