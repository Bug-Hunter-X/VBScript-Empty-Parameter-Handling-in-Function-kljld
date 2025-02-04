Function MyFunction(param1)
  If IsMissing(param1) Or IsEmpty(param1) Then
    ' Handle missing or empty parameter gracefully.
    param1 = ""
  End If
  ' ... rest of the function ...
End Function