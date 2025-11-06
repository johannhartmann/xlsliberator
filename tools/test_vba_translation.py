#!/usr/bin/env python3
"""Test VBA to Python-UNO translation with LLM."""

import os

from loguru import logger

from xlsliberator.vba2py_uno import create_event_handler_stub, translate_vba_to_python


def test_vba_translation() -> None:
    """Test VBA translation with sample macros."""
    if not os.environ.get("ANTHROPIC_API_KEY"):
        logger.error("ANTHROPIC_API_KEY not set")
        return

    # Test case 1: Simple Sub procedure
    vba_simple = """
Sub UpdateCells()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Range("A1").Value = "Hello"
    Range("B1").Value = 42
    Cells(2, 1).Value = "World"
    MsgBox "Updated cells!"
End Sub
"""

    logger.info("=" * 80)
    logger.info("Test 1: Simple Sub Procedure")
    logger.info("=" * 80)
    logger.info(f"Input VBA:\n{vba_simple}")

    result = translate_vba_to_python(vba_simple, use_llm=True)
    logger.info(f"Output Python:\n{result.python_code}")
    logger.info(f"Warnings: {result.warnings}")
    logger.info(f"Unsupported: {result.unsupported_features}")

    # Test case 2: Function with control flow
    vba_function = """
Function CalculateDiscount(price As Double, quantity As Integer) As Double
    Dim discount As Double
    If quantity > 10 Then
        discount = 0.1
    ElseIf quantity > 5 Then
        discount = 0.05
    Else
        discount = 0
    End If
    CalculateDiscount = price * (1 - discount)
End Function
"""

    logger.info("=" * 80)
    logger.info("Test 2: Function with Control Flow")
    logger.info("=" * 80)
    logger.info(f"Input VBA:\n{vba_function}")

    result = translate_vba_to_python(vba_function, use_llm=True)
    logger.info(f"Output Python:\n{result.python_code}")

    # Test case 3: Event handler
    vba_event = """
Private Sub Workbook_Open()
    MsgBox "Welcome to the workbook!"
    Range("A1").Value = "Opened at: " & Now()
    ActiveSheet.Calculate
End Sub
"""

    logger.info("=" * 80)
    logger.info("Test 3: Event Handler")
    logger.info("=" * 80)
    logger.info(f"Input VBA:\n{vba_event}")

    handler = create_event_handler_stub("Workbook_Open", vba_event, use_llm=True)
    logger.info(f"Output Python:\n{handler}")

    # Test case 4: Loop and Range operations
    vba_loop = """
Sub FillRange()
    Dim i As Integer
    For i = 1 To 10
        Cells(i, 1).Value = i * i
        Cells(i, 2).Formula = "=A" & i & "*2"
    Next i
    MsgBox "Range filled!"
End Sub
"""

    logger.info("=" * 80)
    logger.info("Test 4: Loop and Range Operations")
    logger.info("=" * 80)
    logger.info(f"Input VBA:\n{vba_loop}")

    result = translate_vba_to_python(vba_loop, use_llm=True)
    logger.info(f"Output Python:\n{result.python_code}")

    # Test case 5: String functions
    vba_strings = """
Sub ProcessStrings()
    Dim fullName As String
    Dim firstName As String
    Dim lastName As String

    fullName = "John Doe"
    firstName = Left(fullName, 4)
    lastName = Right(fullName, 3)

    Range("A1").Value = UCase(firstName)
    Range("B1").Value = LCase(lastName)
    Range("C1").Value = Len(fullName)
End Sub
"""

    logger.info("=" * 80)
    logger.info("Test 5: String Functions")
    logger.info("=" * 80)
    logger.info(f"Input VBA:\n{vba_strings}")

    result = translate_vba_to_python(vba_strings, use_llm=True)
    logger.info(f"Output Python:\n{result.python_code}")

    logger.success("All VBA translation tests completed!")


if __name__ == "__main__":
    test_vba_translation()
