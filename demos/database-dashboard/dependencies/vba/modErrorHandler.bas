Attribute VB_Name = "modErrorHandler"
' ============================================================
' MinervaDB XL View - Centralized Error Handler Module
' modErrorHandler.bas
' ============================================================
' Provides enterprise-grade error handling, logging, and
' recovery mechanisms for all MinervaDB XL View modules.
' ============================================================
Option Explicit

' ---- Error Log Sheet Name ----
Private Const ERROR_LOG_SHEET As String = "ErrorLog"
Private Const MAX_LOG_ROWS As Long = 10000
Private Const LOG_FILE_PATH As String = "MinervaDB_XL_View_Errors.log"

' ---- Error Severity Levels ----
Public Enum ErrSeverity
    SEV_INFO = 1
    SEV_WARNING = 2
    SEV_ERROR = 3
    SEV_CRITICAL = 4
End Enum

' ---- Error Record Type ----
Private Type ErrorRecord
    Timestamp   As String
    Severity    As String
    Module      As String
    Procedure   As String
    ErrorNumber As Long
    ErrorDesc   As String
    Resolution  As String
End Type

' ============================================================
' Public: Log an error with full context
' ============================================================
Public Sub LogError(ByVal modName As String, _
                    ByVal procName As String, _
                    ByVal errNum As Long, _
                    ByVal errDesc As String, _
                    Optional ByVal severity As ErrSeverity = SEV_ERROR, _
                    Optional ByVal resolution As String = "")

    Dim rec As ErrorRecord
    rec.Timestamp   = Format(Now, "yyyy-mm-dd hh:nn:ss")
    rec.Module      = modName
    rec.Procedure   = procName
    rec.ErrorNumber = errNum
    rec.ErrorDesc   = errDesc
    rec.Resolution  = IIf(resolution = "", "See documentation or contact DBA.", resolution)

    Select Case severity
        Case SEV_INFO:     rec.Severity = "INFO"
        Case SEV_WARNING:  rec.Severity = "WARNING"
        Case SEV_ERROR:    rec.Severity = "ERROR"
        Case SEV_CRITICAL: rec.Severity = "CRITICAL"
    End Select

    ' Write to error log sheet
    WriteToErrorSheet rec

    ' Write to local log file
    WriteToLogFile rec

    ' Display user-facing message for errors and above
    If severity >= SEV_ERROR Then
        ShowErrorDialog rec
    End If

End Sub

' ============================================================
' Public: Handle connection errors
' ============================================================
Public Sub HandleConnectionError(ByVal errNum As Long, ByVal errDesc As String)
    Dim msg As String
    msg = "MinervaDB XL View failed to connect to PostgreSQL." & vbCrLf & _
          "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
          "Possible causes:" & vbCrLf & _
          "  - PostgreSQL server is offline" & vbCrLf & _
          "  - Incorrect host/port in config.ini" & vbCrLf & _
          "  - ODBC driver not installed" & vbCrLf & _
          "  - Firewall blocking port 5432" & vbCrLf & vbCrLf & _
          "Check docs/TROUBLESHOOTING.md for resolution steps."

    MsgBox msg, vbCritical, "MinervaDB XL View - Connection Error"
    LogError "modConnection", "Connect", errNum, errDesc, SEV_CRITICAL
End Sub

' ============================================================
' Public: Handle query errors
' ============================================================
Public Sub HandleQueryError(ByVal queryName As String, _
                             ByVal errNum As Long, _
                             ByVal errDesc As String)
    Dim msg As String
    msg = "MinervaDB XL View query execution failed." & vbCrLf & _
          "Query: " & queryName & vbCrLf & _
          "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
          "The dashboard data may be incomplete. " & _
          "Please refresh or check the query library."

    MsgBox msg, vbExclamation, "MinervaDB XL View - Query Error"
    LogError "modQueryRunner", queryName, errNum, errDesc, SEV_ERROR
End Sub

' ============================================================
' Public: Handle data refresh errors
' ============================================================
Public Sub HandleRefreshError(ByVal sheetName As String, _
                               ByVal errNum As Long, _
                               ByVal errDesc As String)
    LogError "modDataRefresh", sheetName, errNum, errDesc, SEV_WARNING, _
             "Retry refresh manually or check PostgreSQL connectivity."
End Sub

' ============================================================
' Public: Clear the error log sheet
' ============================================================
Public Sub ClearErrorLog()
    Dim ws As Worksheet
    On Error GoTo NoSheet
    Set ws = ThisWorkbook.Sheets(ERROR_LOG_SHEET)
    ws.Rows("2:" & ws.Rows.Count).Delete
    MsgBox "MinervaDB XL View error log cleared.", vbInformation, "MinervaDB XL View"
    Exit Sub
NoSheet:
    MsgBox "Error log sheet not found.", vbExclamation, "MinervaDB XL View"
End Sub

' ============================================================
' Public: Export error log to CSV
' ============================================================
Public Sub ExportErrorLog()
    Dim ws As Worksheet
    Dim fPath As String
    Dim fNum As Integer
    Dim i As Long, lastRow As Long

    On Error GoTo ExportErr

    Set ws = GetOrCreateErrorSheet()
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No errors logged yet.", vbInformation, "MinervaDB XL View"
        Exit Sub
    End If

    fPath = Environ("USERPROFILE") & "\MinervaDB_XL_View_ErrorLog_" & _
            Format(Now, "yyyymmdd_hhnnss") & ".csv"
    fNum = FreeFile

    Open fPath For Output As #fNum
    Print #fNum, "Timestamp,Severity,Module,Procedure,ErrorNumber,Description,Resolution"
    For i = 2 To lastRow
        Print #fNum, _
            """" & ws.Cells(i, 1).Value & """," & _
            """" & ws.Cells(i, 2).Value & """," & _
            """" & ws.Cells(i, 3).Value & """," & _
            """" & ws.Cells(i, 4).Value & """," & _
            ws.Cells(i, 5).Value & "," & _
            """" & ws.Cells(i, 6).Value & """," & _
            """" & ws.Cells(i, 7).Value & """"
    Next i
    Close #fNum

    MsgBox "Error log exported to:" & vbCrLf & fPath, vbInformation, "MinervaDB XL View"
    Exit Sub

ExportErr:
    MsgBox "Failed to export error log: " & Err.Description, vbCritical, "MinervaDB XL View"
End Sub

' ============================================================
' Private: Write error record to log sheet
' ============================================================
Private Sub WriteToErrorSheet(rec As ErrorRecord)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = GetOrCreateErrorSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Trim log if too large
    If nextRow > MAX_LOG_ROWS Then
        ws.Rows("2:101").Delete
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ws.Cells(nextRow, 1).Value = rec.Timestamp
    ws.Cells(nextRow, 2).Value = rec.Severity
    ws.Cells(nextRow, 3).Value = rec.Module
    ws.Cells(nextRow, 4).Value = rec.Procedure
    ws.Cells(nextRow, 5).Value = rec.ErrorNumber
    ws.Cells(nextRow, 6).Value = rec.ErrorDesc
    ws.Cells(nextRow, 7).Value = rec.Resolution

    ' Color-code by severity
    Dim rowColor As Long
    Select Case rec.Severity
        Case "INFO":     rowColor = RGB(200, 230, 200)
        Case "WARNING":  rowColor = RGB(255, 243, 180)
        Case "ERROR":    rowColor = RGB(255, 200, 200)
        Case "CRITICAL": rowColor = RGB(200, 0, 0)
    End Select

    ws.Rows(nextRow).Interior.Color = rowColor
End Sub

' ============================================================
' Private: Write error record to log file
' ============================================================
Private Sub WriteToLogFile(rec As ErrorRecord)
    Dim fNum As Integer
    Dim logPath As String

    On Error GoTo FileErr
    logPath = ThisWorkbook.Path & "\" & LOG_FILE_PATH
    fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, rec.Timestamp & " | " & rec.Severity & " | " & _
                 rec.Module & "." & rec.Procedure & " | " & _
                 "Err " & rec.ErrorNumber & ": " & rec.ErrorDesc
    Close #fNum
    Exit Sub
FileErr:
    ' Silent failure for log file write - don't cascade errors
End Sub

' ============================================================
' Private: Show user-facing error dialog
' ============================================================
Private Sub ShowErrorDialog(rec As ErrorRecord)
    Dim msg As String
    msg = "MinervaDB XL View encountered an error:" & vbCrLf & vbCrLf & _
          "Module:    " & rec.Module & vbCrLf & _
          "Procedure: " & rec.Procedure & vbCrLf & _
          "Error:     " & rec.ErrorNumber & " - " & rec.ErrorDesc & vbCrLf & vbCrLf & _
          "Suggested Resolution:" & vbCrLf & rec.Resolution

    If rec.Severity = "CRITICAL" Then
        MsgBox msg, vbCritical, "MinervaDB XL View - Critical Error"
    Else
        MsgBox msg, vbExclamation, "MinervaDB XL View - Error"
    End If
End Sub

' ============================================================
' Private: Get or create error log worksheet
' ============================================================
Private Function GetOrCreateErrorSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ERROR_LOG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = ERROR_LOG_SHEET

        ' Header row
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Severity"
        ws.Cells(1, 3).Value = "Module"
        ws.Cells(1, 4).Value = "Procedure"
        ws.Cells(1, 5).Value = "Error Number"
        ws.Cells(1, 6).Value = "Description"
        ws.Cells(1, 7).Value = "Resolution"

        ' Format header
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(31, 73, 125)
        ws.Rows(1).Font.Color = RGB(255, 255, 255)
        ws.Columns("A:G").AutoFit
    End If

    Set GetOrCreateErrorSheet = ws
End Function
