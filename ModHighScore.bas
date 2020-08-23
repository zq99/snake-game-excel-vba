Attribute VB_Name = "ModHighScore"
'********************************************************************
'Program: Snake
'Author : Zaid Qureshi
'********************************************************************
Option Explicit

Const cintNoReturn = 0
Const cintMaxLength = 255
Const cstrFileName = "CentipedeScores"
Const cintNone = 0
Const cintStartRow = 4
Const cintIDCol = 4
Const cintExcelNameCol = 5
Const cintScoreCol = 6
Const cstrEmpty = ""
Const cintMaxField = 50
Const cintMaxRecords = 20

Private Type udtHighScore
    UserID As String * 50
    ExcelName As String * 50
    Score As Long
    Date As Date
End Type

Public Function login() As String
    login = Environ("username")
End Function

Public Function AddScore() As Boolean
On Error GoTo errAddHandler
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
    
    AddScore = False
    Set shtGame = Worksheets("Game")

    lngScore = shtGame.Range("Score")
    If lngScore = cintNone Then Exit Function
    intNumOfScores = Application.WorksheetFunction.CountA(Range("ScoreColumn"))
    lngCurrentLowest = Application.WorksheetFunction.Min(Range("ScoreColumn"))
    If (lngScore < lngCurrentLowest) And _
        (intNumOfScores >= cintMaxRecords) _
              Then Exit Function

    lngCurrentHighest = Application.WorksheetFunction.Max(Range("ScoreColumn"))
    strTitle = "High Score!"
    If lngScore > lngCurrentHighest Then
        strMsg = "Congratulations!! " & lngScore & " is a new high score!!!"
        MsgBox strMsg, vbOKOnly + vbExclamation, strTitle
    Else
        strMsg = "Your score of : " & lngScore & Chr(10) & "is good enough for our high score board."
        MsgBox strMsg, vbOKOnly + vbInformation, strTitle
    End If
    
    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    udtHighScore.UserID = IIf(Len(login) > cintMaxField, Mid(login(), 1, cintMaxField - 1), login)
    udtHighScore.ExcelName = IIf(Len(Application.UserName) > cintMaxField, Mid(Application.UserName, 1, cintMaxField - 1), Application.UserName)
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
exitHere:
    Set shtGame = Nothing
    AddScore = True
    Exit Function
errAddHandler:
    AddScore = False
    MsgBox Err.Number & vbTab & Err.Description
    Close intFileNum
    GoTo exitHere
End Function

Public Function GetScores() As Boolean
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
                .Cells(intRow, cintScoreCol).Value = CLng(Trim(udtHighScore.Score))
             End With
             intRow = intRow + 1
          End If
        Loop
    Close intFileNum
    
    With shtScore
        Set rngScore = .Range("ScoreData")
        rngScore.Sort Key1:=.Range("F4"), Order1:=xlDescending _
            , Header:=xlNo, OrderCustom:=1, MatchCase:=False _
            , Orientation:=xlTopToBottom
    End With
    GetScores = True
    
    Exit Function
errGetScoreshandler:
    GetScores = False
    Close intFileNum
End Function


Public Function GetLowestScoreIndex() As Integer
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore
    Dim lngLowestScore As Long
    Dim intCount As Integer
    Dim intindex As Integer
    
    intCount = 1
    intindex = 1
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
                  intindex = intCount
               End If
            End If
        Loop
    Close intFileNum
    GetLowestScoreIndex = intindex
End Function

Public Function NumberOfRecords() As Integer
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim udtHighScore As udtHighScore
    
    intFileNum = FreeFile
    strFileName = ThisWorkbook.Path & "\" & cstrFileName & ".txt"
    Open strFileName For Random As intFileNum Len = Len(udtHighScore)
        NumberOfRecords = LOF(intFileNum) / Len(udtHighScore)
    Close intFileNum
End Function

