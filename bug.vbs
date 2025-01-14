Function MyFunction(param1)
  If IsEmpty(param1) Then
    MyFunction = ""
    Exit Function
  End If
  ' ... rest of the function ...
End Function