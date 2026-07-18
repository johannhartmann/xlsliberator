VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_Open()
    'Purpose: To ensure that the game screen is always selected , when the workbook is open

On Error GoTo errHandler:
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    shtGame.Select
    shtGame.Range(cstrACellThatsNotInTheWay).Select
exitHere:
    Set wkbmacro = Nothing
    Exit Sub
errHandler:
    GoTo exitHere

End Sub
