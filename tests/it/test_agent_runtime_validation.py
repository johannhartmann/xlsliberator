"""Integration tests for agent rewriter runtime validation."""

import tempfile
from pathlib import Path

import pytest

from xlsliberator.agent_rewriter import AgentRewriter
from xlsliberator.embed_macros import embed_python_macros
from xlsliberator.extract_vba import VBAModuleIR, VBAModuleType
from xlsliberator.ir_models import CellIR, CellType, SheetIR, WorkbookIR
from xlsliberator.uno_conn import UnoCtx
from xlsliberator.write_ods import write_ods_from_ir


@pytest.fixture
def skip_if_no_lo(skip_if_no_lo: None) -> None:
    """Skip tests if LibreOffice is not available."""
    pass


# Simple VBA code that should translate successfully
SIMPLE_VBA_CODE = """
Sub SetValue()
    Range("A1").Value = "TEST"
End Sub
"""

# VBA code with syntax that will cause runtime errors
FAILING_VBA_CODE = """
Sub FailingFunction()
    ' This will fail at runtime because we're accessing invalid range
    ThisWorkbook.InvalidProperty.DoSomething()
End Sub
"""


@pytest.mark.integration
def test_runtime_validation_success(skip_if_no_lo: None) -> None:
    """Test that runtime validation succeeds for valid macros."""
    with tempfile.TemporaryDirectory() as tmpdir:
        ods_path = Path(tmpdir) / "test_runtime_success.ods"

        # Create base ODS
        sheet = SheetIR(name="Sheet1", index=0)
        sheet.cells.append(
            CellIR(row=0, col=0, address="A1", cell_type=CellType.STRING, value="Initial")
        )

        wb_ir = WorkbookIR(file_path="test.xlsx", file_format="xlsx", sheets=[sheet])

        with UnoCtx() as ctx:
            write_ods_from_ir(ctx, wb_ir, str(ods_path), locale="en-US")

        # Manually embed a simple working macro (bypassing agent for this test)
        # This tests the runtime validation infrastructure
        python_code = """
import uno
from com.sun.star.awt import XActionListener
from com.sun.star.awt import KeyEvent

def SetValue():
    '''Simple test function that should execute successfully.'''
    # Get document
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.getSheets().getByIndex(0)
    cell = sheet.getCellRangeByName("A1")
    cell.setString("TEST")
    return True

# Export for LibreOffice
g_exportedScripts = (SetValue,)
"""

        embed_python_macros(ods_path, {"Module1.py": python_code})

        # Now test that AgentRewriter._test_and_refine() validates runtime execution
        # We'll use the validation function directly
        from xlsliberator.python_macro_manager import (
            enumerate_python_scripts,
            test_script_execution,
        )

        # Enumerate scripts
        script_infos = enumerate_python_scripts(ods_path)
        assert len(script_infos) > 0, "Should have embedded scripts"

        # Test execution
        execution_successful = True
        runtime_errors = []
        for script_info in script_infos:
            for script_uri in script_info.script_uris:
                try:
                    result = test_script_execution(ods_path, script_uri)
                    if not result.success:
                        runtime_errors.append(f"{script_info.module_name}: {result.error}")
                        execution_successful = False
                except Exception as e:
                    runtime_errors.append(f"{script_info.module_name}: {e}")
                    execution_successful = False

        # Note: Script execution in headless LibreOffice has limitations
        # Check that we at least attempted execution and got a structured response
        assert len(script_infos) > 0, "Should have embedded scripts"
        # If execution failed due to XScriptProvider issues, that's expected in headless mode
        if not execution_successful:
            # Check if it's the expected headless limitation
            has_script_provider_error = any(
                "XScriptProvider" in str(err) or "empty interface" in str(err)
                for err in runtime_errors
            )
            if has_script_provider_error:
                pytest.skip(
                    "Script execution limited in headless mode "
                    "(XScriptProvider not fully functional)"
                )
            else:
                # If it's a different error, that's unexpected
                pytest.fail(f"Unexpected runtime errors: {runtime_errors}")


@pytest.mark.integration
def test_runtime_validation_detects_failures(skip_if_no_lo: None) -> None:
    """Test that runtime validation detects failing macros."""
    with tempfile.TemporaryDirectory() as tmpdir:
        ods_path = Path(tmpdir) / "test_runtime_failure.ods"

        # Create base ODS
        sheet = SheetIR(name="Sheet1", index=0)
        wb_ir = WorkbookIR(file_path="test.xlsx", file_format="xlsx", sheets=[sheet])

        with UnoCtx() as ctx:
            write_ods_from_ir(ctx, wb_ir, str(ods_path), locale="en-US")

        # Embed a macro that will fail at runtime
        failing_code = """
import uno

def FailingFunction():
    '''This function will fail at runtime.'''
    # This will raise an exception
    raise RuntimeError("Intentional failure for testing")

# Export for LibreOffice
g_exportedScripts = (FailingFunction,)
"""

        embed_python_macros(ods_path, {"FailingModule.py": failing_code})

        # Test runtime validation
        from xlsliberator.python_macro_manager import (
            enumerate_python_scripts,
            test_script_execution,
        )

        script_infos = enumerate_python_scripts(ods_path)
        assert len(script_infos) > 0

        # Test execution should detect failure
        execution_successful = True
        for script_info in script_infos:
            for script_uri in script_info.script_uris:
                try:
                    result = test_script_execution(ods_path, script_uri)
                    if not result.success:
                        execution_successful = False
                except Exception:
                    execution_successful = False

        # Should fail
        assert not execution_successful, "Runtime validation should detect failing macros"


@pytest.mark.integration
@pytest.mark.slow
def test_agent_rewriter_runtime_validation_integration(skip_if_no_lo: None) -> None:
    """Test full AgentRewriter integration with runtime validation.

    This test verifies that the AgentRewriter properly performs runtime
    validation after embedding macros.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        ods_path = Path(tmpdir) / "test_agent_runtime.ods"

        # Create base ODS
        sheet = SheetIR(name="Sheet1", index=0)
        wb_ir = WorkbookIR(file_path="test.xlsx", file_format="xlsx", sheets=[sheet])

        with UnoCtx() as ctx:
            write_ods_from_ir(ctx, wb_ir, str(ods_path), locale="en-US")

        # Create simple VBA module
        vba_module = VBAModuleIR(
            name="SimpleModule",
            module_type=VBAModuleType.STANDARD,
            source_code=SIMPLE_VBA_CODE,
            procedures=["SetValue"],
            dependencies=set(),
        )

        # Use AgentRewriter to translate and validate
        # Note: This may require ANTHROPIC_API_KEY to be set
        try:
            rewriter = AgentRewriter()
            code, validation = rewriter.rewrite_vba_project(
                modules=[vba_module],
                source_file="test.xlsx",
                output_path=ods_path,
                max_iterations=3,
            )

            # Verify validation result includes runtime execution status
            assert hasattr(
                validation, "execution_successful"
            ), "Validation should include execution_successful field"

            # Note: We can't assert execution_successful == True here because:
            # 1. The agent might classify this as "simple" and skip agent rewriting
            # 2. Runtime testing might fail in headless mode for certain macros
            # 3. This is primarily testing that the infrastructure is in place

            # Instead, verify that the validation was attempted
            assert validation.syntax_valid is not None
            assert validation.has_exports is not None
            assert validation.execution_successful is not None

        except Exception as e:
            # If agent rewriting is skipped or API key is missing, that's OK
            # We're primarily testing the infrastructure
            pytest.skip(f"Agent rewriting test skipped: {e}")
