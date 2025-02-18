Function MyFunction(param1, param2)
  On Error Resume Next
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Clear ' Clear previous errors
    result = "Parameters cannot be empty"
  Else
    ' ... function logic to execute if parameters are valid...
    result = param1 + param2 'Example
  End If
  On Error GoTo 0
  MyFunction = result
End Function