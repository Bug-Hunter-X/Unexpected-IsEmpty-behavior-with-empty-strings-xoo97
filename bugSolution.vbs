Function MyFunction(param1)
  If VarType(param1) = vbVariant Then
    If Len(param1) = 0 Then
      MyFunction = ""
      Exit Function
    ElseIf IsEmpty(param1) Then
      MyFunction = ""
      Exit Function
    End If
  End If
  ' ... rest of the function ...
End Function