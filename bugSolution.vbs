Function f(a)
  Dim i, prev, curr, next
  prev = 1
  curr = 1
  If a = 1 Then
    f = 1
  ElseIf a = 2 Then
    f = 1
  Else
    For i = 3 To a
      next = prev + curr
      prev = curr
      curr = next
    Next
    f = curr
  End If
End Function
MsgBox f(10)