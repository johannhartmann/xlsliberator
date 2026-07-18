Attribute VB_Name = "ModHighScore"
'*********************************************************
'*Purpose: to handle the high scores facility
'*********************************************************
Option Explicit

'Declare constants
Private Const cintNoReturn As Integer = 0
Private Const cintMaxLength As Integer = 255
Private Const cstrFileName As String = "ExceltrisHighScores"
Private Const cintNone As Integer = 0
Private Const cintStartRow As Integer = 4
Private Const cintIDCol As Integer = 4
Private Const cintExcelNameCol As Integer = 5
Private Const cintTotalRowCol As Integer = 6
Private Const cintScoreCol As Integer = 7
Private Const cstrEmpty As String = ""
Private Const cintMaxField As Integer = 50
Private Const cintMaxRecords As Integer = 20

Private Type udtHighScore
    UserID As String * 50
    ExcelName As String * 50
    Score As Long
    Rows As Long
    Date As Date
End Type

Public Function login() As String
     login = Environ("username")
End Function

Public Function AddScore() As Boolean
    'Purpose: Add a highscore to the file
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore
    Dim shtGame As Worksheet
    Dim intNumberOfRecords As Integer
    Dim intFileRow As Integer
    Dim lngScore As Long
    Dim lngCurrentLowest As Long
    Dim lngCurrentHighest As Long
    Dim strMsg As String
    Dim strTitle As String
    Dim intNumOfScores As Integer

    On Error GoTo errAddHandler
    AddScore = False
    Set shtGame = Worksheets("Game")

    'Check if user's score is worthy of the score table..
    lngScore = shtGame.Range("Score")
    If lngScore = cintNone Then Exit Function 'no action if zero
    intNumOfScores = Application.WorksheetFunction.CountA(Range("ScoreColumn"))
    lngCurrentLowest = Application.WorksheetFunction.Min(Range("ScoreColumn"))
    If (lngScore < lngCurrentLowest) And _
        (intNumOfScores >= cintMaxRecords) _
              Then Exit Function 'no action if its not higher then the lowest on the board

    'Inform user of their score and where it sits in the universe..
    lngCurrentHighest = Application.WorksheetFunction.Max(Range("ScoreColumn"))
    strTitle = "High Score!"
    If lngScore > lngCurrentHighest Then
        strMsg = "Congratulations!! " & lngScore & " is a new high score!!!"
        MsgBox strMsg, vbOKOnly + vbExclamation, strTitle
    Else
        strMsg = "Your score of : " & lngScore & Chr(10) & "is good enough for our high score board."
        MsgBox strMsg, vbOKOnly + vbInformation, strTitle
    End If

    'Process a high score...
    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    udtHighScore.UserID = IIf(Len(login) > cintMaxField, Mid(login(), 1, cintMaxField - 1), login)
    udtHighScore.ExcelName = IIf(Len(Application.UserName) > cintMaxField, Mid(Application.UserName, 1, cintMaxField - 1), _
                                            Application.UserName)
    udtHighScore.Rows = shtGame.Range("totalRows")
    udtHighScore.Score = lngScore
    udtHighScore.Date = Now()

    If NumberOfRecords < cintMaxRecords Then
        Open strFileName For Random As intFileNum Len = Len(udtHighScore)
            intNumberOfRecords = LOF(intFileNum) / Len(udtHighScore)
            intFileRow = intNumberOfRecords + 1
            Put intFileNum, intFileRow, udtHighScore
        Close intFileNum
    Else
        intFileRow = GetLowestScoreIndex
        Open strFileName For Random As intFileNum Len = Len(udtHighScore)
            Put intFileNum, intFileRow, udtHighScore
        Close intFileNum
    End If

    Set shtGame = Nothing
    AddScore = True
    Exit Function

errAddHandler:
    AddScore = False
    MsgBox Err.Number & vbTab & Err.Description
    Close intFileNum
    Set shtGame = Nothing
End Function

Public Function GetScores() As Boolean
    'Purpose: Get Scores from file
On Error GoTo errGetScoreshandler

    Dim intCount As Integer
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore
    Dim shtScore As Worksheet
    Dim rngScore As Range
    Dim intRow As Integer

    GetScores = False
    intCount = cintNone
    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    Set shtScore = ThisWorkbook.Worksheets("Score")
    shtScore.Range("ScoreData").Cells.ClearContents
    intRow = cintStartRow
    Open strFileName For Random As intFileNum Len = Len(udtHighScore)
        Do While (Not EOF(intFileNum))
            intCount = intCount + 1
            Get intFileNum, intCount, udtHighScore
            If Trim(Application.WorksheetFunction.Clean(udtHighScore.UserID)) <> cstrEmpty Then
               With shtScore
                  .Cells(intRow, cintIDCol).Value = Trim(CStr(udtHighScore.UserID))
                  .Cells(intRow, cintExcelNameCol).Value = Trim(CStr(udtHighScore.ExcelName))
                  .Cells(intRow, cintTotalRowCol).Value = CLng(Trim(udtHighScore.Rows))
                  .Cells(intRow, cintScoreCol).Value = CLng(Trim(udtHighScore.Score))
               End With
               intRow = intRow + 1
            End If
        Loop
    Close intFileNum

    With shtScore
        Set rngScore = .Range("ScoreData")
        rngScore.Sort Key1:=.Range("G4"), Order1:=xlDescending, Key2:=.Range("F4") _
            , Order2:=xlDescending, Header:=xlNo, OrderCustom:=1, MatchCase:=False _
            , Orientation:=xlTopToBottom
    End With
    GetScores = True

    Exit Function
errGetScoreshandler:
    GetScores = False
    Close intFileNum
End Function


Public Function GetLowestScoreIndex() As Integer
    'Purpose: to get the position of the lowest score in the random access file.
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore
    Dim lngLowestScore As Long
    Dim intCount As Integer
    Dim intIndex As Integer

    intCount = 1
    intIndex = 1
    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    Open strFileName For Random As intFileNum Len = Len(udtHighScore)
        Get intFileNum, 1, udtHighScore
        lngLowestScore = udtHighScore.Score
        Do While (Not EOF(intFileNum))
            intCount = intCount + 1
            Get intFileNum, intCount, udtHighScore
            If Trim(Application.WorksheetFunction.Clean(udtHighScore.UserID)) <> cstrEmpty Then
               If udtHighScore.Score < lngLowestScore Then
                    lngLowestScore = udtHighScore.Score
                    intIndex = intCount
               End If
            End If
        Loop
Close intFileNum
    GetLowestScoreIndex = intIndex
End Function

Public Function NumberOfRecords() As Integer
    'Purpose: to get the number of records in the file
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore

    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    Open strFileName For Random As intFileNum Len = Len(udtHighScore)
        NumberOfRecords = LOF(intFileNum) / Len(udtHighScore)
    Close intFileNum
End Function
