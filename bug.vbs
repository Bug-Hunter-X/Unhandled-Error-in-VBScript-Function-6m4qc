Function MyFunction(param1)
  'Some code here that might throw an error
  If Err.Number <> 0 Then
    'Handle the error
  End If
End Function