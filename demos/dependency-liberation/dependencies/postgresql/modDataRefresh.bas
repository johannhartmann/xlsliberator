Attribute VB_Name = "modDataRefresh"
' ============================================================
' MinervaDB XL View - Automated Data Refresh Module
' modDataRefresh.bas
' ============================================================
' Handles scheduled and on-demand data refresh for all
' MinervaDB XL View dashboard sheets.
' ============================================================
Option Explicit

' ---- Refresh Configuration ----
Private Const DEFAULT_INTERVAL_MIN As Integer = 15
Private Const MAX_RETRY_ATTEMPTS As Integer = 3
Private Const RETRY_DELAY_SEC As Integer = 5

' ---- Module-Level Variables ----
Private m_RefreshInterval As Integer
Private m_AutoRefreshEnabled As Boolean
Private m_LastRefreshTime As Date
Private m_RefreshCount As Long

' ---- Sheet Name Constants ----
Private Const SHEET_SALES As String = "Sales_Dashboard"
Private Const SHEET_FINANCE As String = "Finance_Dashboard"
Private Const SHEET_HR As String = "HR_Dashboard"
Private Const SHEET_OPS As String = "Operations_Dashboard"
Private Const SHEET_EXEC As String = "Executive_Summary"

' ============================================================
' Public: Initialize the refresh module
' ============================================================
Public Sub InitializeRefresh(Optional ByVal intervalMinutes As Integer = DEFAULT_INTERVAL_MIN)
    m_RefreshInterval = intervalMinutes
    m_AutoRefreshEnabled = False
    m_LastRefreshTime = #1/1/1900#
    m_RefreshCount = 0

    ' Load interval from config if available
    On Error Resume Next
    Dim cfgInterval As String
    cfgInterval = modConnection.GetConfigValue("refresh", "interval_minutes")
    If cfgInterval <> "" Then m_RefreshInterval = CInt(cfgInterval)
    On Error GoTo 0

    modErrorHandler.LogError "modDataRefresh", "InitializeRefresh", 0, _
        "MinervaDB XL View refresh module initialized. Interval: " & m_RefreshInterval & " min", _
        SEV_INFO
End Sub

' ============================================================
' Public: Refresh all dashboard sheets
' ============================================================
Public Sub RefreshAllDashboards()
    Dim startTime As Single
    Dim conn As Object
    Dim attempt As Integer

    startTime = Timer
    Application.StatusBar = "MinervaDB XL View: Refreshing all dashboards..."
    Application.ScreenUpdating = False

    ' Attempt connection with retry
    For attempt = 1 To MAX_RETRY_ATTEMPTS
        Set conn = modConnection.GetConnection()
        If Not conn Is Nothing Then Exit For

        If attempt < MAX_RETRY_ATTEMPTS Then
            Application.StatusBar = "MinervaDB XL View: Connection attempt " & attempt & " failed. Retrying..."
            Application.Wait Now + TimeSerial(0, 0, RETRY_DELAY_SEC)
        End If
    Next attempt

    If conn Is Nothing Then
        modErrorHandler.HandleConnectionError -1, "All " & MAX_RETRY_ATTEMPTS & " connection attempts failed."
        GoTo Cleanup
    End If

    ' Refresh each dashboard
    RefreshSalesDashboard conn
    RefreshFinanceDashboard conn
    RefreshHRDashboard conn
    RefreshOperationsDashboard conn
    RefreshExecutiveSummary conn

    ' Update metadata
    m_LastRefreshTime = Now
    m_RefreshCount = m_RefreshCount + 1
    UpdateRefreshMetadata

    modErrorHandler.LogError "modDataRefresh", "RefreshAllDashboards", 0, _
        "All dashboards refreshed in " & Format(Timer - startTime, "0.00") & "s", SEV_INFO

Cleanup:
    If Not conn Is Nothing Then
        modConnection.CloseConnection conn
    End If
    Application.ScreenUpdating = True
    Application.StatusBar = "MinervaDB XL View: Last refresh " & Format(Now, "hh:nn:ss")
End Sub

' ============================================================
' Public: Refresh a single named dashboard
' ============================================================
Public Sub RefreshDashboard(ByVal dashboardName As String)
    Dim conn As Object
    Set conn = modConnection.GetConnection()

    If conn Is Nothing Then
        modErrorHandler.HandleConnectionError -1, "Cannot refresh " & dashboardName & " - no connection."
        Exit Sub
    End If

    Application.StatusBar = "MinervaDB XL View: Refreshing " & dashboardName & "..."

    Select Case dashboardName
        Case SHEET_SALES:   RefreshSalesDashboard conn
        Case SHEET_FINANCE: RefreshFinanceDashboard conn
        Case SHEET_HR:      RefreshHRDashboard conn
        Case SHEET_OPS:     RefreshOperationsDashboard conn
        Case SHEET_EXEC:    RefreshExecutiveSummary conn
        Case Else:
            modErrorHandler.LogError "modDataRefresh", "RefreshDashboard", 0, _
                "Unknown dashboard: " & dashboardName, SEV_WARNING
    End Select

    modConnection.CloseConnection conn
    Application.StatusBar = "MinervaDB XL View: " & dashboardName & " refreshed at " & Format(Now, "hh:nn:ss")
End Sub

' ============================================================
' Public: Start auto-refresh timer
' ============================================================
Public Sub StartAutoRefresh()
    m_AutoRefreshEnabled = True
    ScheduleNextRefresh
    modErrorHandler.LogError "modDataRefresh", "StartAutoRefresh", 0, _
        "Auto-refresh started. Interval: " & m_RefreshInterval & " minutes.", SEV_INFO
    MsgBox "MinervaDB XL View: Auto-refresh started." & vbCrLf & _
           "Interval: every " & m_RefreshInterval & " minutes.", _
           vbInformation, "MinervaDB XL View"
End Sub

' ============================================================
' Public: Stop auto-refresh timer
' ============================================================
Public Sub StopAutoRefresh()
    On Error Resume Next
    Application.OnTime m_LastRefreshTime + TimeSerial(0, m_RefreshInterval, 0), _
        "modDataRefresh.AutoRefreshCallback", Schedule:=False
    On Error GoTo 0
    m_AutoRefreshEnabled = False
    Application.StatusBar = "MinervaDB XL View: Auto-refresh stopped."
    MsgBox "MinervaDB XL View: Auto-refresh has been stopped.", _
           vbInformation, "MinervaDB XL View"
End Sub

' ============================================================
' Public: Callback procedure for OnTime scheduler
' ============================================================
Public Sub AutoRefreshCallback()
    If Not m_AutoRefreshEnabled Then Exit Sub
    RefreshAllDashboards
    ScheduleNextRefresh
End Sub

' ============================================================
' Public: Get refresh status info
' ============================================================
Public Function GetRefreshStatus() As String
    Dim status As String
    status = "MinervaDB XL View Refresh Status" & vbCrLf & _
             "=================================" & vbCrLf & _
             "Auto-Refresh: " & IIf(m_AutoRefreshEnabled, "ON", "OFF") & vbCrLf & _
             "Interval: " & m_RefreshInterval & " minutes" & vbCrLf & _
             "Total Refreshes: " & m_RefreshCount & vbCrLf & _
             "Last Refresh: " & IIf(m_LastRefreshTime = #1/1/1900#, "Never", _
                                    Format(m_LastRefreshTime, "yyyy-mm-dd hh:nn:ss"))
    GetRefreshStatus = status
End Function

' ============================================================
' Private: Refresh Sales Dashboard
' ============================================================
Private Sub RefreshSalesDashboard(ByVal conn As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_SALES)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Application.StatusBar = "MinervaDB XL View: Refreshing Sales Dashboard..."

    ' Load KPIs
    Dim rsKPI As Object
    Set rsKPI = modQueryRunner.ExecuteQuery(conn, "SELECT * FROM vw_sales_dashboard WHERE report_date >= CURRENT_DATE - INTERVAL '30 days' ORDER BY report_date DESC")
    If Not rsKPI Is Nothing Then
        modQueryRunner.ResultSetToSheet rsKPI, ws, "A2", True
        rsKPI.Close
    End If

    ' Update timestamp
    ws.Range("Z1").Value = "Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Sub

ErrHandler:
    modErrorHandler.HandleRefreshError SHEET_SALES, Err.Number, Err.Description
End Sub

' ============================================================
' Private: Refresh Finance Dashboard
' ============================================================
Private Sub RefreshFinanceDashboard(ByVal conn As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_FINANCE)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Application.StatusBar = "MinervaDB XL View: Refreshing Finance Dashboard..."
    Dim rs As Object
    Set rs = modQueryRunner.ExecuteQuery(conn, "SELECT * FROM vw_finance_summary ORDER BY period DESC LIMIT 24")
    If Not rs Is Nothing Then
        modQueryRunner.ResultSetToSheet rs, ws, "A2", True
        rs.Close
    End If
    ws.Range("Z1").Value = "Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Sub
ErrHandler:
    modErrorHandler.HandleRefreshError SHEET_FINANCE, Err.Number, Err.Description
End Sub

' ============================================================
' Private: Refresh HR Dashboard
' ============================================================
Private Sub RefreshHRDashboard(ByVal conn As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_HR)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Application.StatusBar = "MinervaDB XL View: Refreshing HR Dashboard..."
    Dim rs As Object
    Set rs = modQueryRunner.ExecuteQuery(conn, "SELECT * FROM vw_hr_metrics ORDER BY metric_date DESC LIMIT 365")
    If Not rs Is Nothing Then
        modQueryRunner.ResultSetToSheet rs, ws, "A2", True
        rs.Close
    End If
    ws.Range("Z1").Value = "Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Sub
ErrHandler:
    modErrorHandler.HandleRefreshError SHEET_HR, Err.Number, Err.Description
End Sub

' ============================================================
' Private: Refresh Operations Dashboard
' ============================================================
Private Sub RefreshOperationsDashboard(ByVal conn As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_OPS)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Application.StatusBar = "MinervaDB XL View: Refreshing Operations Dashboard..."
    Dim rs As Object
    Set rs = modQueryRunner.ExecuteQuery(conn, "SELECT * FROM vw_operations_kpis ORDER BY metric_date DESC LIMIT 90")
    If Not rs Is Nothing Then
        modQueryRunner.ResultSetToSheet rs, ws, "A2", True
        rs.Close
    End If
    ws.Range("Z1").Value = "Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Sub
ErrHandler:
    modErrorHandler.HandleRefreshError SHEET_OPS, Err.Number, Err.Description
End Sub

' ============================================================
' Private: Refresh Executive Summary
' ============================================================
Private Sub RefreshExecutiveSummary(ByVal conn As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_EXEC)
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub

    Application.StatusBar = "MinervaDB XL View: Refreshing Executive Summary..."
    Dim rs As Object
    Set rs = modQueryRunner.ExecuteQuery(conn, "SELECT * FROM vw_executive_kpis ORDER BY kpi_date DESC LIMIT 1")
    If Not rs Is Nothing Then
        modQueryRunner.ResultSetToSheet rs, ws, "B2", False
        rs.Close
    End If
    ws.Range("Z1").Value = "Refreshed: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    Exit Sub
ErrHandler:
    modErrorHandler.HandleRefreshError SHEET_EXEC, Err.Number, Err.Description
End Sub

' ============================================================
' Private: Update refresh metadata on a dedicated sheet
' ============================================================
Private Sub UpdateRefreshMetadata()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RefreshLog")
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = m_RefreshCount
    ws.Cells(nextRow, 3).Value = "SUCCESS"
End Sub

' ============================================================
' Private: Schedule the next auto-refresh
' ============================================================
Private Sub ScheduleNextRefresh()
    If Not m_AutoRefreshEnabled Then Exit Sub
    Dim nextTime As Date
    nextTime = Now + TimeSerial(0, m_RefreshInterval, 0)
    Application.OnTime nextTime, "modDataRefresh.AutoRefreshCallback"
End Sub
