'''
'the below script contains the lsit of buttons added and its fucntion
'''
Private Sub CommandButton1_Click()
'show fail subjects button
    Dim subject As Range, xcell As Range
    Set subject = Range("b2:i6")
  For Each xcell In subject
    If xcell.Value < 40 Then
        With xcell
            .Font.Bold = True
            .Interior.Color = vbRed
        End With
    End If
  Next xcell
End Sub

Private Sub CommandButton2_Click()
    'reset button
    Dim subject As Range
    Set subject = Range("a2:i6")
    With subject
        .Font.Bold = False
        .Interior.ColorIndex = 0
        .Borders.LineStyle = xlna
    End With
End Sub

Private Sub CommandButton3_Click()
'show failed students
    Dim resultcells As Range, xcell As Range
    Dim failstud As Range
    Set resultcells = Range("j2:j6")
    For Each xcell In resultcells
        If xcell.Value = "Fail" Then
          Set failstud = Cells(xcell.Row, 1)
          With failstud
            .Font.Bold = True
            .Interior.Color = vbYellow
            .Borders.Color = vbBlack
          End With
        End If
    Next xcell
End Sub


'''
'Codes for the sub and functions which are executed to calculate values.
'''
Function GetTotalScores(rowno As Integer)
    'to get the total score of student
    If rowno = Empty Then GetTotalScores = "NA"
    Dim studMarks As Range
    Set studMarks = Range("b" & rowno & ":i" & rowno)
    totalMarks = 0
    For i = 1 To studMarks.Count
       totalMarks = totalMarks + studMarks.Item(i).Value
    Next i
    
    GetTotalScores = totalMarks
End Function

Function StudResult(obj1 As Range) As String
'returns either the student is Pass or Fail
    If obj1 Is Nothing Then Exit Function

    Dim failCnt As Integer
    Dim cell As Range
    failCnt = 0
    
    For Each cell In obj1
        If cell.Value < 40 Then
            failCnt = failCnt + 1
        End If
    Next cell
    
    If failCnt = 0 Then
        StudResult = "Pass"
    Else
        StudResult = "Fail"
    End If
End Function

Function GetGradeLetter(perCell As Range)
    Dim per As Double, result As String
    per = perCell.Value
    
  resultval = Cells(perCell.Row, perCell.Column - 2).Value
  GetGradeLetter = "F"
  If resultval = "Fail" Then Exit Function
    Select Case per
        Case Is < 0.6
            GetGradeLetter = "F"
        Case Is < 0.7
             GetGradeLetter = "D"
        Case Is < 0.8
           GetGradeLetter = "C"
        Case Is < 0.9
            GetGradeLetter = "B"
        Case Else
            GetGradeLetter = "A"
    End Select
End Function


