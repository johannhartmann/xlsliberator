Attribute VB_Name = "ModGame"
'************************************************************************
'Program   :  EXCELTRIS
'Author    :  Zaid Qureshi
'Version   :  1.0
'************************************************************************

Option Explicit
Option Base 1

Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Const KeyPressed As Integer = -32767
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" _
        (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'shape definitions
Private Const cintNumOfSquares As Integer = 4
Private Const cintShapeLLeft As Integer = 1
Private Const cintShapeLRight As Integer = 9
Private Const cintShapeT As Integer = 2
Private Const cintShapeRectangle As Integer = 3
Private Const cintShapeCube As Integer = 4
Private Const cintShapeSmallRect As Integer = 5
Private Const cintShapeSingle As Integer = 6
Private Const cintShapeTwistLeft As Integer = 7
Private Const cintShapeTwistRight As Integer = 8
Private Const cintRed As Integer = 3
Private Const cintYellow As Integer = 6
Private Const cintBlue As Integer = 41
Private Const cintGreen As Integer = 4
Private Const cintPink As Integer = 40
Private Const cintOrange As Integer = 45
Private Const cintPurple As Integer = 54
Private Const cintGrey As Integer = 48

Private Const cintMaxShapes As Integer = 9
Private Const cintFirstShape As Integer = 1
Private Const clngStartDelay As Integer = 160

'For directions..
Private Const cintDown As Integer = 1
Private Const cintDownQuick As Integer = 2
Private Const cintLeft As Integer = -1
Private Const cintRight As Integer = 1
Private Const cintStraight As Integer = 0

'For grid definition..
Private Const cintBottomRow As Integer = 31
Private Const cintTopRow As Integer = 4
Private Const cintFirstColumn As Integer = 13
Private Const cintLastColumn As Integer = 27
Private Const cintNone As Integer = 0
Public Const cstrACellThatsNotInTheWay As String = "a32"

'For points..
Private Const cintBonusPoint  As Integer = 1
Private Const cintBrickDropPoints As Integer = 5
Private Const cintPointsForARow As Integer = 100
Private Const cintSingleRow  As Integer = 1

'General..
Private Const cintUserInterrupt As Integer = 18
Private Const cstrClearRange As String = "M4:AA31"
Private Const cstrGameTitle As String = "EXCELTRIS"

'For shape status
Private Const cintNormal As Integer = 1
Private Const cintSideLeft As Integer = 2
Private Const cintUpSideDown As Integer = 3
Private Const cintSideRight As Integer = 4

'For Sound effects
Private Const cstrCollapseSound As String = "\collapse.wav"
Private Const cstrBrickLandSound As String = "\land.wav"

'Declare modular variables
Private mrngGrid As Range
Private mrngTop As Range
Private mlngDelay As Long
Private mshtGame As Worksheet
Private marrCurRange() As Range
Private mrngShape As Range
Private mintColor As Integer
Private marrBase() As Integer
Private mblnGameInProgress As Boolean
Private mintNew As Integer
Private mintStatus As Integer
Private mrngPileOfBricks As Range

Public Sub DrawBrick()

    Dim intIndex As Integer
    Dim rngStartBlock As Range
    Static intPreview As Integer

    With mshtGame
        Set mrngGrid = Union(.Range("leftboundary"), .Range("bottomboundary"), .Range("rightboundary"))
        Set mrngTop = .Range("topboundary")
        Randomize
        If intPreview = cintNone Then
             intPreview = Int((cintMaxShapes - cintFirstShape + 1) * Rnd + cintFirstShape)
             ShowPreview intPreview
             mintNew = Int((cintMaxShapes - cintFirstShape + 1) * Rnd + cintFirstShape)
        Else
             mintNew = intPreview
             intPreview = Int((cintMaxShapes - cintFirstShape + 1) * Rnd + cintFirstShape)
             ShowPreview intPreview
        End If
        Select Case mintNew
        Case cintShapeLLeft
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("s6")
            Set marrCurRange(4) = .Range("t6")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintRed
            ReDim marrBase(1 To 2)
            marrBase(1) = 3
            marrBase(2) = 4
        Case cintShapeLRight
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("s6")
            Set marrCurRange(4) = .Range("r6")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintRed
            ReDim marrBase(1 To 2)
            marrBase(1) = 3
            marrBase(2) = 4
        Case cintShapeRectangle
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("s6")
            Set marrCurRange(4) = .Range("s7")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            Set mrngShape = marrCurRange(1)
            mintColor = cintYellow
            ReDim marrBase(1)
            marrBase(1) = 4
        Case cintShapeT
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("s6")
            Set marrCurRange(4) = .Range("t5")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintGreen
            ReDim marrBase(2)
            marrBase(1) = 3
            marrBase(2) = 4
        Case cintShapeCube
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("t4")
            Set marrCurRange(4) = .Range("t5")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintBlue
            ReDim marrBase(1 To 2)
            marrBase(1) = 2
            marrBase(2) = 4
        Case cintShapeSmallRect
            ReDim marrCurRange(1 To 2)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2))
            mintColor = cintGrey
            ReDim marrBase(1)
            marrBase(1) = 2
        Case cintShapeSingle
            ReDim marrCurRange(1)
            Set marrCurRange(1) = .Range("s4")
            Set mrngShape = marrCurRange(1)
            mintColor = cintOrange
            ReDim marrBase(1)
            marrBase(1) = 1
        Case cintShapeTwistLeft
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("t5")
            Set marrCurRange(4) = .Range("t6")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintPurple
            ReDim marrBase(1 To 2)
            marrBase(1) = 2
            marrBase(2) = 4
        Case cintShapeTwistRight
            ReDim marrCurRange(1 To 4)
            Set marrCurRange(1) = .Range("s4")
            Set marrCurRange(2) = .Range("s5")
            Set marrCurRange(3) = .Range("r5")
            Set marrCurRange(4) = .Range("r6")
            Set mrngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
            mintColor = cintPurple
            ReDim marrBase(1 To 2)
            marrBase(1) = 2
            marrBase(2) = 4
        End Select
        With mrngShape.Interior
            .ColorIndex = mintColor
            .Pattern = xlSolid
        End With
        mintStatus = cintNormal
    End With
    Set rngStartBlock = Nothing
End Sub

Public Sub Run()
On Error GoTo errStartStophandler:
    Set mshtGame = Worksheets("game")
    With mshtGame
        If Trim(.Shapes("cmdStartStop").TextFrame.Characters.Text) = "Start" Then
           .Shapes("cmdStartStop").TextFrame.Characters.Text = "Stop"
           .Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintRed
            Call StartGame
        Else
           .Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
           .Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
            End
        End If
    End With
    Exit Sub
errStartStophandler:
    ResetBoard
    mblnGameInProgress = False
    mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
End Sub

Private Sub StartGame()

    Dim intIndex As Integer
    Dim arrDestRange() As Range
    Dim blnBrickNeeded As Boolean
    Dim blnGameOver As Boolean
    Dim strMsg As String
    Dim strTitle As String
    Dim rngDestShape As Range
    Dim rngCurrentShape As Range
    Dim blnBrickLanded  As Boolean
    Dim blnMove As Boolean
    Dim rngFinal As Range

    On Error GoTo errMoveHandler

    Application.DisplayAlerts = False
    Application.EnableCancelKey = xlErrorHandler

    If mshtGame Is Nothing Then _
        Set mshtGame = Worksheets("game")
    ResetBoard
    mlngDelay = clngStartDelay
    blnBrickNeeded = True
    blnGameOver = False
    mblnGameInProgress = True
    Set mrngPileOfBricks = Nothing

    Do
         DoEvents
         Sleep mlngDelay

         If blnBrickNeeded Then
            DrawBrick
            blnBrickNeeded = False
         End If
         Set rngDestShape = Nothing

         blnMove = True
         Select Case KeyPressed
         ReDim arrDestRange(LBound(marrCurRange()) To UBound(marrCurRange()))
         Case GetAsyncKeyState(vbKeyLeft)
             For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                   Set arrDestRange(intIndex) = marrCurRange(intIndex).Offset(cintDown, cintLeft)
                   If rngDestShape Is Nothing Then Set rngDestShape = arrDestRange(intIndex) Else _
                        Set rngDestShape = Union(rngDestShape, arrDestRange(intIndex))
             Next
         Case GetAsyncKeyState(vbKeyRight)
             For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                   Set arrDestRange(intIndex) = marrCurRange(intIndex).Offset(cintDown, cintRight)
                   If rngDestShape Is Nothing Then Set rngDestShape = arrDestRange(intIndex) Else _
                        Set rngDestShape = Union(rngDestShape, arrDestRange(intIndex))
             Next
         Case GetAsyncKeyState(vbKeyDown)
             For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                   Set arrDestRange(intIndex) = marrCurRange(intIndex).Offset(cintDownQuick, cintStraight)
                   If rngDestShape Is Nothing Then Set rngDestShape = arrDestRange(intIndex) Else _
                        Set rngDestShape = Union(rngDestShape, arrDestRange(intIndex))
             Next
             'extra point if your quick
             Call updateScore(cintBonusPoint)
         Case GetAsyncKeyState(vbKeyUp), GetAsyncKeyState(vbKeyControl)
             Call Rotate
             blnMove = False
         Case Else
             For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                   Set arrDestRange(intIndex) = marrCurRange(intIndex).Offset(cintDown, cintStraight)
                   If rngDestShape Is Nothing Then Set rngDestShape = arrDestRange(intIndex) Else _
                        Set rngDestShape = Union(rngDestShape, arrDestRange(intIndex))
             Next
         End Select

         Range(cstrACellThatsNotInTheWay).Select 'stop the cursor from moving everywhere

         If blnMove Then '#1
           If Not TestForClashes(rngDestShape) Then '#3
               Set rngCurrentShape = Nothing
               'Group together the current shape cell ranges..
               For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                   If rngCurrentShape Is Nothing Then Set rngCurrentShape = marrCurRange(intIndex) Else _
                        Set rngCurrentShape = Union(rngCurrentShape, marrCurRange(intIndex))
                       'set the current shape to equal the new shape
                        Set marrCurRange(intIndex) = arrDestRange(intIndex)
               Next

               rngCurrentShape.Interior.ColorIndex = xlNone 'remove color from old position
               rngDestShape.Interior.ColorIndex = mintColor 'add color to new position

               'if a brick has reached solid ground check the base references for the shape to see if
               'underneath them there is solid ground , if so, then brick has landed
               blnBrickLanded = False
               For intIndex = LBound(marrBase()) To UBound(marrBase())
                   If arrDestRange(marrBase(intIndex)).Offset(cintDown, cintStraight) _
                                    .Interior.ColorIndex <> xlNone Then
                            blnBrickLanded = True
                   End If
               Next
           End If '#3

           If blnBrickLanded Then '#2
                PlaySound ActiveWorkbook.Path & cstrBrickLandSound, &H1, &H1
                'increment score for a brick land
                Call updateScore(cintBrickDropPoints)
                blnBrickNeeded = True
                For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
                    If rngFinal Is Nothing Then Set rngFinal = marrCurRange(intIndex) Else _
                        Set rngFinal = Union(rngFinal, marrCurRange(intIndex))
                Next
                If mrngPileOfBricks Is Nothing Then Set mrngPileOfBricks = rngFinal Else _
                    Set mrngPileOfBricks = Union(mrngPileOfBricks, rngFinal)
                Call checkForCompleteLines
                'check if the top has been reached
                If Not Intersect(Range("topboundary"), rngFinal) Is Nothing Then
                    blnGameOver = True
                    Exit Do
                End If
           End If '#2
         End If '#1
    Loop

    If blnGameOver Then
        MsgBox "Game Over!!", vbOKOnly, cstrGameTitle
        Call AddScore
        Call GetScores
        Worksheets("Score").Select
        Call ResetBoard
        mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
        mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
        mblnGameInProgress = False
        End
    End If

exitHere:
    Set rngDestShape = Nothing
    Exit Sub
errMoveHandler:
    Select Case Err
    Case cintUserInterrupt
        strMsg = "Are you sure you want to cancel?"
        strTitle = cstrGameTitle
        If MsgBox(strMsg, vbYesNo + vbExclamation, strTitle) = vbYes Then
            MsgBox "Game Halted"
            mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
            mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
            mblnGameInProgress = False
            End
        Else
            Resume
        End If
    Case Else
        MsgBox Err.Number & vbTab & Err.Description
        mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
        mshtGame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
        mblnGameInProgress = False
    End Select
    GoTo exitHere
End Sub

Public Sub ResetBoard()
    Dim rngGame As Range
    Set mshtGame = Worksheets("game")
    Set rngGame = mshtGame.Range(cstrClearRange)
    rngGame.Cells.ClearFormats
    rngGame.Cells.ClearContents
    mshtGame.Range("previewblock").Cells.Clear
    mshtGame.Range("totalrows").Value = cintNone
    mshtGame.Range("Score").Value = cintNone
    Set rngGame = Nothing
End Sub

Public Function checkForCompleteLines() As Boolean

    Dim intRow As Integer
    Dim intCol As Integer
    Dim blnCompleteRow As Boolean
    Dim rngAreaAboveCompleteRow As Range
    Dim rngNewLocation As Range
    Dim rngDest As Range

    checkForCompleteLines = False
    intRow = cintBottomRow
    Do Until intRow = cintTopRow
        If mshtGame.Cells(intRow, cintFirstColumn).Interior.ColorIndex <> xlNone Then
            blnCompleteRow = True
            For intCol = cintFirstColumn To cintLastColumn
                If mshtGame.Cells(intRow, intCol).Interior.ColorIndex = xlNone Then
                    blnCompleteRow = False
                    Exit For
                End If
            Next
            If blnCompleteRow Then
                'select range above the complete row
                Set rngAreaAboveCompleteRow = mshtGame.Range(mshtGame.Cells(cintTopRow, cintFirstColumn), _
                                      mshtGame.Cells(intRow - 1, cintLastColumn))
                'select newlocation (block displacement of 1 row)..
                Set rngNewLocation = mshtGame.Range(mshtGame.Cells(cintTopRow + 1, cintFirstColumn), mshtGame.Cells(intRow, cintLastColumn))

                'move data down, over complete row
                rngAreaAboveCompleteRow.Cut Destination:=rngNewLocation

                Set rngAreaAboveCompleteRow = Nothing
                Set rngNewLocation = Nothing
                intRow = intRow + 1 'this to make sure that after a delete the row is repeated
                'since after the above line is executed then intRow - 1 makes
                'neutralize intRow
                Call updateScore(cintPointsForARow, cintSingleRow)
                PlaySound ActiveWorkbook.Path & cstrCollapseSound, &H1, &H1
            End If
        End If
        intRow = intRow - 1
    Loop
    checkForCompleteLines = True

    Set rngAreaAboveCompleteRow = Nothing
    Set rngNewLocation = Nothing
    Set rngDest = Nothing

End Function

Public Sub updateScore(Optional ByVal intPoints, Optional ByVal intRow)
    Dim lngRows As Long

    If Not IsMissing(intPoints) Then mshtGame.Range("Score").Value = mshtGame.Range("Score") + intPoints
    If Not IsMissing(intRow) Then mshtGame.Range("totalRows").Value = mshtGame.Range("totalrows").Value + intRow

    'adjust speed according the users progress the more rows the faster it gets
    lngRows = mshtGame.Range("totalrows").Value
    Select Case lngRows
      Case cintNone To 9
             mlngDelay = 160
      Case 10 To 19
             mlngDelay = 145
      Case 20 To 29
             mlngDelay = 120
      Case 30 To 39
             mlngDelay = 110
      Case 40 To 49
             mlngDelay = 100
      Case 50 To 59
             mlngDelay = 90
      Case 60 To 69
             mlngDelay = 70
      Case 70 To 79
             mlngDelay = 60
      Case 80 To 99
             mlngDelay = 50
      Case Else
             mlngDelay = 40
    End Select

End Sub

Public Sub About_Game()
    Dim strMsg As String
    Dim strMsgtitle As String

    strMsgtitle = "ABOUT"
    strMsg = cstrGameTitle & " - Programmed by Zaid Qureshi"
    MsgBox strMsg, vbInformation, strMsgtitle
End Sub

Public Sub ViewHighScores()
    Dim wkbmacro As Workbook
    Dim shtScore As Worksheet
    Set wkbmacro = ThisWorkbook
    Set shtScore = wkbmacro.Worksheets("Score")
    If mblnGameInProgress = False Then
        GetScores
        shtScore.Select
    End If
    Set wkbmacro = Nothing
    Set shtScore = Nothing
End Sub

Public Sub ViewGameScreen()
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    shtGame.Select
    shtGame.Range(cstrACellThatsNotInTheWay).Select
    Set wkbmacro = Nothing
    Set shtGame = Nothing
End Sub

Public Sub ShowPreview(ByVal intPreviewShape)
'Purpose: to draw the next brick, so that the user has

    Dim rngPreview As Range
    Dim arrPreviewRange() As Range
    Dim rngPreviewShape As Range
    Dim intColor As Integer

    With mshtGame
        Set rngPreview = .Range("previewblock")
        rngPreview.Clear
        Set rngPreview = Nothing
        Select Case intPreviewShape
        Case cintShapeLLeft
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("g20")
            Set rngPreviewRange(4) = .Range("h20")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintRed
        Case cintShapeLRight
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("g20")
            Set rngPreviewRange(4) = .Range("f20")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintRed
        Case cintShapeRectangle
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("g20")
            Set rngPreviewRange(4) = .Range("g21")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintYellow
        Case cintShapeT
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("g20")
            Set rngPreviewRange(4) = .Range("h19")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintGreen
        Case cintShapeCube
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("h18")
            Set rngPreviewRange(4) = .Range("h19")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintBlue
        Case cintShapeSmallRect
            ReDim rngPreviewRange(1 To 2)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2))
            intColor = cintGrey
        Case cintShapeSingle
            ReDim rngPreviewRange(1)
            Set rngPreviewRange(1) = .Range("g20")
            Set rngPreviewShape = rngPreviewRange(1)
            intColor = cintOrange
        Case cintShapeTwistLeft
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("g18")
            Set rngPreviewRange(2) = .Range("g19")
            Set rngPreviewRange(3) = .Range("h19")
            Set rngPreviewRange(4) = .Range("h20")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintPurple
        Case cintShapeTwistRight
            ReDim rngPreviewRange(1 To 4)
            Set rngPreviewRange(1) = .Range("h18")
            Set rngPreviewRange(2) = .Range("h19")
            Set rngPreviewRange(3) = .Range("g19")
            Set rngPreviewRange(4) = .Range("g20")
            Set rngPreviewShape = Union(rngPreviewRange(1), rngPreviewRange(2), rngPreviewRange(3), rngPreviewRange(4))
            intColor = cintPurple
        End Select
        With rngPreviewShape.Interior
            .ColorIndex = intColor
            .Pattern = xlSolid
        End With
    End With

    Set rngPreview = Nothing
    Set rngPreviewShape = Nothing

End Sub


Public Sub Rotate()
'Purpose:  To allow the user to rotate an object..

    Dim rngShape As Range
    Dim arrOriginalShapeRange() As Range ' to hold the original shape
    Dim arrOriginalBasePoints() As Integer ' to hold the original base points
    Dim intIndex As Integer

    'store the position of the original range in case we need to reset things..
    ReDim arrOriginalShapeRange(LBound(marrCurRange()) To UBound(marrCurRange()))
    For intIndex = LBound(marrCurRange()) To UBound(marrCurRange())
        Set arrOriginalShapeRange(intIndex) = marrCurRange(intIndex)
    Next

    'Store the original base points..
    ReDim arrOriginalBasePoints(LBound(marrBase()) To UBound(marrBase()))
    For intIndex = LBound(marrBase()) To UBound(marrBase())
        arrOriginalBasePoints(intIndex) = marrBase(intIndex)
    Next

    Select Case mintNew
    Case cintShapeLLeft

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(-2, 0)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 3
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(0, -2)
             mintStatus = cintUpSideDown
             ReDim marrBase(1 To 2)
             marrBase(1) = 1
             marrBase(2) = 4
        Case cintUpSideDown
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(2, 0)
             mintStatus = cintSideRight
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 4
        Case cintSideRight
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(0, 2)
             mintStatus = cintNormal
             ReDim marrBase(1 To 2)
             marrBase(1) = 3
             marrBase(2) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintRed

    Case cintShapeLRight

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(0, 2)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 4
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(-2, 0)
             mintStatus = cintUpSideDown
             ReDim marrBase(1 To 2)
             marrBase(1) = 1
             marrBase(2) = 4
        Case cintUpSideDown
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(0, -2)
             mintStatus = cintSideRight
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 3
        Case cintSideRight
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(2, 0)
             mintStatus = cintNormal
             ReDim marrBase(1 To 2)
             marrBase(1) = 3
             marrBase(2) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintRed

    Case cintShapeRectangle

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(2, -2)
             Set marrCurRange(2) = marrCurRange(2).Offset(1, -1)
             Set marrCurRange(3) = marrCurRange(3).Offset(0, 0)
             Set marrCurRange(4) = marrCurRange(4).Offset(-1, 1)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 4)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 3
             marrBase(4) = 4
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(-2, 2)
             Set marrCurRange(2) = marrCurRange(2).Offset(-1, 1)
             Set marrCurRange(3) = marrCurRange(3).Offset(0, 0)
             Set marrCurRange(4) = marrCurRange(4).Offset(1, -1)
             mintStatus = cintNormal
             ReDim marrBase(1)
             marrBase(1) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintYellow

    Case cintShapeT

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(-1, -1)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 3
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(1, -1)
             mintStatus = cintUpSideDown
             ReDim marrBase(1 To 2)
             marrBase(1) = 1
             marrBase(2) = 4
        Case cintUpSideDown
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(1, 1)
             mintStatus = cintSideRight
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 3
             marrBase(3) = 4
        Case cintSideRight
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(-1, 1)
             mintStatus = cintNormal
             ReDim marrBase(1 To 2)
             marrBase(1) = 3
             marrBase(2) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintGreen

    Case cintShapeCube

       'Not worth bothering about rotation
       'since shape never changes when rotated...

    Case cintShapeSmallRect

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 2)
             marrBase(1) = 1
             marrBase(2) = 2
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             mintStatus = cintNormal
             ReDim marrBase(1)
             marrBase(1) = 2
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2))
        mintColor = cintGrey

    Case cintShapeSingle

       'Not worth bothering about rotation
       'since shape never changes when rotated...

    Case cintShapeTwistLeft

       'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(1, -1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(-1, -1)
             Set marrCurRange(4) = marrCurRange(4).Offset(-2, 0)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 2
             marrBase(3) = 4
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(-1, 1)
             Set marrCurRange(2) = marrCurRange(2).Offset(0, 0)
             Set marrCurRange(3) = marrCurRange(3).Offset(1, 1)
             Set marrCurRange(4) = marrCurRange(4).Offset(2, 0)
             mintStatus = cintNormal
             ReDim marrBase(1 To 2)
             marrBase(1) = 2
             marrBase(2) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintPurple

    Case cintShapeTwistRight

        'Clear current shape on screen
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        rngShape.Clear
        'Decide what position current shape was in..
        Select Case mintStatus
        Case cintNormal
             Set marrCurRange(1) = marrCurRange(1).Offset(0, -2)
             Set marrCurRange(2) = marrCurRange(2).Offset(-1, -1)
             Set marrCurRange(3) = marrCurRange(3).Offset(0, 0)
             Set marrCurRange(4) = marrCurRange(4).Offset(-1, 1)
             mintStatus = cintSideLeft
             ReDim marrBase(1 To 3)
             marrBase(1) = 1
             marrBase(2) = 3
             marrBase(3) = 4
        Case cintSideLeft
             Set marrCurRange(1) = marrCurRange(1).Offset(0, 2)
             Set marrCurRange(2) = marrCurRange(2).Offset(1, 1)
             Set marrCurRange(3) = marrCurRange(3).Offset(0, 0)
             Set marrCurRange(4) = marrCurRange(4).Offset(1, -1)
             mintStatus = cintNormal
             ReDim marrBase(1 To 2)
             marrBase(1) = 2
             marrBase(2) = 4
        End Select
        Set rngShape = Union(marrCurRange(1), marrCurRange(2), marrCurRange(3), marrCurRange(4))
        mintColor = cintPurple

    End Select

    If (mintNew <> cintShapeSingle) And (mintNew <> cintShapeCube) Then
        'Check if the new shape will clash with any of the barriers
        'or any other shape , if so we cannot allow rotation
        If TestForClashes(rngShape) Then
            'reset shape to original position...
            ReDim marrCurRange(LBound(arrOriginalShapeRange()) To UBound(arrOriginalShapeRange()))
            For intIndex = LBound(arrOriginalShapeRange()) To UBound(arrOriginalShapeRange())
                Set marrCurRange(intIndex) = arrOriginalShapeRange(intIndex)
            Next
            'reset base points to original set
            ReDim marrBase(LBound(arrOriginalBasePoints()) To UBound(arrOriginalBasePoints()))
            For intIndex = LBound(arrOriginalBasePoints()) To UBound(arrOriginalBasePoints())
                marrBase(intIndex) = arrOriginalBasePoints(intIndex)
            Next
        Else
        'redraw new shape..
        With rngShape.Interior
            .ColorIndex = mintColor
            .Pattern = xlSolid
             End With
        End If
    End If

    Set rngShape = Nothing

End Sub

Public Function TestForClashes(ByVal rngTest As Range) As Boolean
    'Purpose: to test if a shape will clash with (1) the barrier , (2) any other shapes

    TestForClashes = False
    If Not Intersect(mrngGrid, rngTest) Is Nothing Then TestForClashes = True

    'Make sure that the user's selection does not involve overwriting
    'any existing bricks that are already on the screen..
    If Not mrngPileOfBricks Is Nothing Then
        If Not Intersect(mrngPileOfBricks, rngTest) Is Nothing Then
            TestForClashes = True
        End If
    End If

End Function
