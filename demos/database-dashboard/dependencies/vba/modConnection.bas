Attribute VB_Name = "modConnection"
'==============================================================================
' modConnection.bas
' PostgreSQL Connection Manager Module
' Excel-PostgreSQL-Dashboard Repository
' Author: Shiv Iyer | MinervaDB | ChistaDATA
' License: MIT
'==============================================================================

Option Explicit

'--- ADODB Connection Object (module-level singleton) ---
Private m_Conn      As Object   ' ADODB.Connection
Private m_ConnStr   As String
Private m_IsOpen    As Boolean

'--- Default connection parameters ---
Private Const DEF_PROVIDER  As String = "MSDASQL"   ' ODBC Provider
Private Const DEF_PORT      As Long   = 5432
Private Const DEF_TIMEOUT   As Long   = 30
Private Const DEF_SCHEMA    As String = "public"

'==============================================================================
' Public API
'==============================================================================

'------------------------------------------------------------------------------
' PG_Connect
' Establishes a connection to PostgreSQL using supplied parameters.
' Returns True on success, False on failure.
'------------------------------------------------------------------------------
Public Function PG_Connect( _
        ByVal sHost     As String, _
        ByVal sDB       As String, _
        ByVal sUser     As String, _
        ByVal sPassword As String, _
        Optional ByVal lPort    As Long   = 5432, _
        Optional ByVal sSchema  As String = "public", _
        Optional ByVal bSSL     As Boolean = True) As Boolean

    On Error GoTo ErrHandler

    ' Disconnect any existing session
    If m_IsOpen Then PG_Disconnect

    Set m_Conn = CreateObject("ADODB.Connection")

    ' Build ODBC connection string
    m_ConnStr = BuildConnectionString(sHost, sDB, sUser, sPassword, lPort, sSchema, bSSL)

    m_Conn.ConnectionTimeout = DEF_TIMEOUT
    m_Conn.CommandTimeout    = DEF_TIMEOUT
    m_Conn.Open m_ConnStr

    m_IsOpen     = True
    PG_Connect   = True

    modErrorHandler.LogInfo "PG_Connect", "Connected to " & sDB & "@" & sHost & ":" & lPort
    Exit Function

ErrHandler:
    m_IsOpen   = False
    PG_Connect = False
    modErrorHandler.LogError "PG_Connect", Err.Number, Err.Description
End Function

'------------------------------------------------------------------------------
' PG_ConnectDSN
' Connects using a pre-configured ODBC DSN (simplest method for end users).
'------------------------------------------------------------------------------
Public Function PG_ConnectDSN( _
        ByVal sDSN      As String, _
        ByVal sUser     As String, _
        ByVal sPassword As String) As Boolean

    On Error GoTo ErrHandler

    If m_IsOpen Then PG_Disconnect

    Set m_Conn = CreateObject("ADODB.Connection")
    m_ConnStr  = "DSN=" & sDSN & ";UID=" & sUser & ";PWD=" & sPassword & ";"
    m_Conn.ConnectionTimeout = DEF_TIMEOUT
    m_Conn.Open m_ConnStr
    m_IsOpen       = True
    PG_ConnectDSN  = True

    modErrorHandler.LogInfo "PG_ConnectDSN", "Connected via DSN: " & sDSN
    Exit Function

ErrHandler:
    m_IsOpen      = False
    PG_ConnectDSN = False
    modErrorHandler.LogError "PG_ConnectDSN", Err.Number, Err.Description
End Function

'------------------------------------------------------------------------------
' PG_Disconnect
' Closes the active PostgreSQL connection gracefully.
'------------------------------------------------------------------------------
Public Sub PG_Disconnect()
    On Error Resume Next
    If Not m_Conn Is Nothing Then
        If m_Conn.State = 1 Then m_Conn.Close  ' adStateOpen = 1
    End If
    Set m_Conn = Nothing
    m_IsOpen   = False
    modErrorHandler.LogInfo "PG_Disconnect", "Connection closed."
End Sub

'------------------------------------------------------------------------------
' PG_IsConnected
' Returns True if the connection is currently open.
'------------------------------------------------------------------------------
Public Function PG_IsConnected() As Boolean
    PG_IsConnected = m_IsOpen And Not (m_Conn Is Nothing)
    If PG_IsConnected Then
        If m_Conn.State <> 1 Then
            m_IsOpen       = False
            PG_IsConnected = False
        End If
    End If
End Function

'------------------------------------------------------------------------------
' PG_GetConnection
' Returns the active ADODB.Connection object (for use in other modules).
'------------------------------------------------------------------------------
Public Function PG_GetConnection() As Object
    If PG_IsConnected() Then
        Set PG_GetConnection = m_Conn
    Else
        Set PG_GetConnection = Nothing
    End If
End Function

'------------------------------------------------------------------------------
' PG_TestConnection
' Runs a lightweight query to verify the connection is alive.
'------------------------------------------------------------------------------
Public Function PG_TestConnection() As Boolean
    Dim rs As Object  ' ADODB.Recordset
    On Error GoTo ErrHandler

    If Not PG_IsConnected() Then
        PG_TestConnection = False
        Exit Function
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT 1 AS ping", m_Conn, 0, 1  ' adOpenForwardOnly, adLockReadOnly

    PG_TestConnection = (Not rs.EOF)
    rs.Close
    Set rs = Nothing
    Exit Function

ErrHandler:
    PG_TestConnection = False
    modErrorHandler.LogError "PG_TestConnection", Err.Number, Err.Description
End Function

'------------------------------------------------------------------------------
' PG_GetServerVersion
' Returns the PostgreSQL server version string.
'------------------------------------------------------------------------------
Public Function PG_GetServerVersion() As String
    Dim rs As Object
    On Error GoTo ErrHandler

    If Not PG_IsConnected() Then
        PG_GetServerVersion = "Not connected"
        Exit Function
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT version()", m_Conn, 0, 1
    PG_GetServerVersion = rs.Fields(0).Value
    rs.Close
    Set rs = Nothing
    Exit Function

ErrHandler:
    PG_GetServerVersion = "Error: " & Err.Description
    modErrorHandler.LogError "PG_GetServerVersion", Err.Number, Err.Description
End Function

'==============================================================================
' Connection Dialog (User Interface)
'==============================================================================

'------------------------------------------------------------------------------
' ShowConnectionDialog
' Displays a user-friendly input dialog to collect connection parameters.
' Stores credentials securely using modSecurity.
'------------------------------------------------------------------------------
Public Sub ShowConnectionDialog()
    Dim sHost     As String
    Dim sDB       As String
    Dim sUser     As String
    Dim sPassword As String
    Dim lPort     As Long

    ' Read saved settings from config sheet (if exists)
    sHost = GetConfigValue("PG_HOST", "localhost")
    sDB   = GetConfigValue("PG_DATABASE", "")
    sUser = GetConfigValue("PG_USERNAME", "excel_dashboard_ro")
    lPort = CLng(GetConfigValue("PG_PORT", "5432"))

    ' Prompt user for each field
    sHost = InputBox("PostgreSQL Host:", "Connection Settings", sHost)
    If sHost = "" Then Exit Sub

    sDB = InputBox("Database Name:", "Connection Settings", sDB)
    If sDB = "" Then Exit Sub

    sUser = InputBox("Username:", "Connection Settings", sUser)
    If sUser = "" Then Exit Sub

    sPassword = InputBox("Password (will not be stored in plain text):", _
                         "Connection Settings", "")

    ' Attempt connection
    If PG_Connect(sHost, sDB, sUser, sPassword, lPort) Then
        MsgBox "Connected successfully to PostgreSQL!" & vbCrLf & _
               PG_GetServerVersion(), vbInformation, "Connection Success"
        ' Save non-sensitive settings
        SaveConfigValue "PG_HOST",     sHost
        SaveConfigValue "PG_DATABASE", sDB
        SaveConfigValue "PG_USERNAME", sUser
        SaveConfigValue "PG_PORT",     CStr(lPort)
    Else
        MsgBox "Connection failed. Please check your settings and try again.", _
               vbCritical, "Connection Error"
    End If
End Sub

'==============================================================================
' Private Helpers
'==============================================================================

Private Function BuildConnectionString( _
        sHost As String, sDB As String, _
        sUser As String, sPassword As String, _
        lPort As Long, sSchema As String, _
        bSSL As Boolean) As String

    Dim cs As String
    cs = "Driver={PostgreSQL Unicode(x64)};"
    cs = cs & "Server=" & sHost & ";"
    cs = cs & "Port=" & lPort & ";"
    cs = cs & "Database=" & sDB & ";"
    cs = cs & "Uid=" & sUser & ";"
    cs = cs & "Pwd=" & sPassword & ";"
    cs = cs & "Options=search_path%3D" & sSchema & ";"
    If bSSL Then cs = cs & "SSLmode=require;"
    BuildConnectionString = cs
End Function

Private Function GetConfigValue(sKey As String, sDefault As String) As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    If ws Is Nothing Then
        GetConfigValue = sDefault
        Exit Function
    End If

    Dim cell As Range
    Set cell = ws.Columns(1).Find(sKey, LookAt:=xlWhole)
    If cell Is Nothing Then
        GetConfigValue = sDefault
    Else
        GetConfigValue = CStr(cell.Offset(0, 1).Value)
    End If
End Function

Private Sub SaveConfigValue(sKey As String, sValue As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Config")
    If ws Is Nothing Then Exit Sub

    Dim cell As Range
    Set cell = ws.Columns(1).Find(sKey, LookAt:=xlWhole)
    If cell Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(lastRow, 1).Value = sKey
        ws.Cells(lastRow, 2).Value = sValue
    Else
        cell.Offset(0, 1).Value = sValue
    End If
End Sub
