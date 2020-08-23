Attribute VB_Name = "ModSnake"
'********************************************************************
'Program: Snake
'Author : Zaid Qureshi
'********************************************************************
Option Explicit
Option Base 1
Public Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Const KeyPressed As Integer = -32767
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const cintLeft = 1
Private Const cintRight = 2
Private Const cintUp = 3
Private Const cintDown = 4
Private Const cstrGameTitle = "SNAKE"
Private Const cintStartlength = 4
Private Const cintNone = 0
Private Const cintFirst = 1
Private Const cintRed = 3
Private Const cintUserInterrupt = 18
Private Const cintGreen = 4
Private Const cstrACellThatsNotInTheWay = "a1"
Private Const cintMaxColumn = 30
Private Const cintMinColumn = 17
Private Const cintMaxRow = 38
Private Const cintMinRow = 3
Private Const cstrHiddenCell = "a32"
Private Const cintMoveLeft = -1
Private Const cintMoveRight = 1
Private Const cintMoveUp = -1
Private Const cintMoveDown = 1
Private Const cintAdvancedDelay = 80
Private Const cintNormalDelay = 120
Private Const cintBeginnerDelay = 140
Private Const cintAdvanced = 1
Private Const cintNormal = 2
Private Const cintBeginner = 3
Private Const cintAdvancedPoints = 100
Private Const cintNormalPoints = 75
Private Const cintBeginnerPoints = 50

Private mrngSnake() As Range
Private mrngSnakeShape As Range
Private mshtgame As Worksheet
Private mrngNext As Range
Private mrngTail As Range
Private mrngTemp As Range
Private mrngBoundary As Range
Private mbolGameInProgress As Boolean
Private mintDirection As Integer
Private mintNumberEaten As Integer

Public Sub Run()
On Error GoTo errorhandler
    Set mshtgame = ThisWorkbook.Worksheets("game")
    With mshtgame
        If Trim(.Shapes("cmdStartStop").TextFrame.Characters.Text) = "Start" Then
            .Shapes("cmdStartStop").TextFrame.Characters.Text = "Stop"
            .Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintRed
             Call Move
        Else
            .Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
            .Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
             End
        End If
    End With
exitHere:
    Set mshtgame = Nothing
    Exit Sub
errorhandler:
    MsgBox Err.Number & vbTab & Err.Description
    mbolGameInProgress = False
    GoTo exitHere
End Sub

Public Sub Move()
On Error GoTo errMoveHandler
    Dim strTitle As String
    Dim strMsg As String
    Dim intindex As Integer
    Dim bolGameOver As Boolean
    Dim strEndMessage As String
    Dim bolFoodNeeded As Boolean
    Dim intNumberOfParts As Integer
    Dim rngSecondFromLast As Range
    Dim rngLast As Range
    Dim lngDelay As Long
    
    Application.DisplayAlerts = False
    Application.EnableCancelKey = xlErrorHandler
    
    Call Clear
    ReDim mrngSnake(cintFirst To cintStartlength)
    With mshtgame
       Set mrngSnake(1) = .Range("r21")
       Set mrngSnake(2) = .Range("s21")
       Set mrngSnake(3) = .Range("t21")
       Set mrngSnake(4) = .Range("u21")
       .Range("aq11").Interior.ColorIndex = cintGreen
    End With
    Set mrngSnakeShape = Nothing
    For intindex = LBound(mrngSnake()) To UBound(mrngSnake())
        If mrngSnakeShape Is Nothing Then Set mrngSnakeShape = mrngSnake(intindex) Else _
            Set mrngSnakeShape = Union(mrngSnakeShape, mrngSnake(intindex))
    Next
    mrngSnakeShape.Interior.ColorIndex = cintRed
    
    Set mrngNext = mrngSnake(UBound(mrngSnake()))
    mintDirection = cintRight
    mbolGameInProgress = True
    bolGameOver = False
    bolFoodNeeded = False
    mintNumberEaten = cintNone
    Set mrngBoundary = mshtgame.Range("boundary")
    Do
         DoEvents
         lngDelay = Delay * Decrease
         Sleep lngDelay
         
         If bolFoodNeeded Then
            SupplyFood
            bolFoodNeeded = False
         End If
         
         Select Case KeyPressed
         Case GetAsyncKeyState(vbKeyLeft)
            mintDirection = cintLeft
         Case GetAsyncKeyState(vbKeyRight)
            mintDirection = cintRight
         Case GetAsyncKeyState(vbKeyDown)
            mintDirection = cintDown
         Case GetAsyncKeyState(vbKeyUp)
            mintDirection = cintUp
         Case Else
            mintDirection = mintDirection
         End Select
         
         Select Case mintDirection
         Case cintLeft
            Set mrngNext = mrngNext.Offset(cintNone, cintMoveLeft)
         Case cintRight
            Set mrngNext = mrngNext.Offset(cintNone, cintMoveRight)
         Case cintDown
            Set mrngNext = mrngNext.Offset(cintMoveDown, cintNone)
         Case cintUp
            Set mrngNext = mrngNext.Offset(cintMoveUp, cintNone)
         End Select

         mshtgame.Range(cstrHiddenCell).Select
         
         If mrngNext.Interior.ColorIndex = cintGreen Then
            mintNumberEaten = mintNumberEaten + 1
            updateScore
            bolFoodNeeded = True
            intNumberOfParts = UBound(mrngSnake)
            ReDim Preserve mrngSnake(cintFirst To intNumberOfParts + 1)
            For intindex = UBound(mrngSnake) To LBound(mrngSnake) Step -1
                If intindex = LBound(mrngSnake) Then
                    Set rngSecondFromLast = mrngSnake(intindex + 1)
                    Set rngLast = mrngSnake(intindex)
                    If rngSecondFromLast = rngLast.Offset(cintMoveUp, cintNone) Then _
                        Set mrngSnake(intindex) = rngLast.Offset(cintMoveDown, cintNone)
                    If rngSecondFromLast = rngLast.Offset(cintMoveDown, cintNone) Then _
                        Set mrngSnake(intindex) = rngLast.Offset(cintMoveUp, cintNone)
                    If rngSecondFromLast = rngLast.Offset(cintNone, cintMoveLeft) Then _
                        Set mrngSnake(intindex) = rngLast.Offset(cintNone, cintMoveRight)
                    If rngSecondFromLast = rngLast.Offset(cintNone, cintMoveRight) Then _
                        Set mrngSnake(intindex) = rngLast.Offset(cintNone, cintMoveLeft)
                Else
                    Set mrngSnake(intindex) = mrngSnake(intindex - 1)
                End If
            Next
         End If
         
         If Not Intersect(mrngNext, mshtgame.Range("boundary")) Is Nothing Then
            strEndMessage = "You have hit the wall!!"
            bolGameOver = True
            Exit Do
         End If
         
         If Not Intersect(mrngNext, mrngSnakeShape) Is Nothing Then
            strEndMessage = "You tried to eat yourself!"
            bolGameOver = True
            Exit Do
         End If
         
         Set mrngSnakeShape = Nothing
         If Intersect(mrngBoundary, mrngSnake(LBound(mrngSnake))) Is Nothing Then _
            mrngSnake(LBound(mrngSnake)).Interior.ColorIndex = xlNone
         For intindex = cintFirst To UBound(mrngSnake())
            If intindex = UBound(mrngSnake()) Then
               Set mrngSnake(intindex) = mrngNext
            Else
               Set mrngSnake(intindex) = mrngSnake(intindex + 1)
            End If
            If mrngSnakeShape Is Nothing Then Set mrngSnakeShape = mrngSnake(intindex) Else _
               Set mrngSnakeShape = Union(mrngSnakeShape, mrngSnake(intindex))
         Next
         mrngNext.Interior.ColorIndex = cintRed
    Loop
    
    If bolGameOver Then
        MsgBox strEndMessage, vbExclamation, "Game Over"
        AddScore
        GetScores
        Worksheets("Score").Select
        mbolGameInProgress = False
        mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
        mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
        Clear
    End If
    
exitHere:
    Set rngSecondFromLast = Nothing
    Set rngLast = Nothing
    Exit Sub
errMoveHandler:
    Select Case Err
    Case cintUserInterrupt
        strMsg = "Are you sure you want to cancel?"
        strTitle = cstrGameTitle
        If MsgBox(strMsg, vbYesNo + vbExclamation, strTitle) = vbYes Then
            MsgBox "Game Halted"
            mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
            mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
            mbolGameInProgress = False
            End
        Else
            Resume
        End If
    Case Else
        MsgBox Err.Number & vbTab & Err.Description
        mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Text = "Start"
        mshtgame.Shapes("cmdStartStop").TextFrame.Characters.Font.ColorIndex = cintGreen
        mbolGameInProgress = False
    End Select
End Sub

Public Sub SupplyFood()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    Do
        intRow = Int((cintMaxRow - cintMinRow + 1) * Rnd + cintMinRow)
        intCol = Int((cintMaxColumn - cintMinColumn + 1) * Rnd + cintMinColumn)
        If Intersect(shtGame.Cells(intRow, intCol), mrngSnakeShape) Is Nothing Then
            Exit Do
        End If
    Loop
    With shtGame.Cells(intRow, intCol)
        .Interior.ColorIndex = cintGreen
    End With
    Set wkbmacro = Nothing
    Set shtGame = Nothing
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
    If mbolGameInProgress = False Then
        Call GetScores
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

Public Sub Clear()
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    shtGame.Range("lair").Cells.Clear
    shtGame.Range("Score").Value = cintNone
    With shtGame.Range("r21:u21")
        .Interior.ColorIndex = cintRed
    End With
    shtGame.Range("aq11").Interior.ColorIndex = cintGreen
    Set wkbmacro = Nothing
    Set shtGame = Nothing
End Sub

Public Sub updateScore()
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    shtGame.Range("Score").Value = shtGame.Range("Score").Value + points
    Set wkbmacro = Nothing
    Set shtGame = Nothing
End Sub

Public Function points() As Integer
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    Dim intLevel As Integer
    
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    With shtGame
        intLevel = .Range("UserLevelSelection")
        Select Case intLevel
        Case cintAdvanced
           points = cintAdvancedPoints
        Case cintNormal
           points = cintNormalPoints
        Case cintBeginner
           points = cintBeginnerPoints
        End Select
    End With
    Set wkbmacro = Nothing
    Set shtGame = Nothing
End Function

Public Function Delay() As Long
    Dim wkbmacro As Workbook
    Dim shtGame As Worksheet
    Dim intLevel As Integer
    
    Set wkbmacro = ThisWorkbook
    Set shtGame = wkbmacro.Worksheets("game")
    With shtGame
        intLevel = .Range("UserLevelSelection")
        Select Case intLevel
        Case cintAdvanced
            Delay = cintAdvancedDelay
        Case cintNormal
            Delay = cintNormalDelay
        Case cintBeginner
            Delay = cintBeginnerDelay
        End Select
    End With
    Set wkbmacro = Nothing
    Set shtGame = Nothing
End Function

Private Function Decrease() As Long
     Select Case mintNumberEaten
     Case cintNone To 9
        Decrease = 1
     Case 10 To 19
        Decrease = 0.95
     Case 20 To 29
        Decrease = 0.9
     Case 30 To 39
        Decrease = 0.85
     Case 40 To 49
        Decrease = 0.8
     Case 50 To 59
        Decrease = 0.75
     Case 60 To 69
        Decrease = 0.7
     Case 70 To 79
        Decrease = 0.65
     Case 80 To 89
        Decrease = 0.6
     Case 90 To 99
        Decrease = 0.5
     Case Else
        Decrease = 0.4
     End Select
End Function
