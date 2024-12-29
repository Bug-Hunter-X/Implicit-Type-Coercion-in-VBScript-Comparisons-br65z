Function f(x)
  On Error Resume Next 'Handle potential type mismatch errors
  If IsNumeric(x) Then
    If x < 0 Then
      f = -1
    ElseIf x = 0 Then
      f = 0
    Else
      f = 1
    End If
  Else
    f = -2 ' Or handle non-numeric input appropriately
  End If
  On Error GoTo 0
End Function

MsgBox f(-1)
MsgBox f("test")
MsgBox f(True) 