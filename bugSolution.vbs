Function MyFunction(param1)
  On Error Resume Next 'Enable error handling
  'Some code here that might throw an error
  If Err.Number <> 0 Then
    'Log the error (e.g., to a file or event log)
    'Example:  Write error details to a file
    Dim fso, file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile("error.log", True)
    file.WriteLine "Error Number: " & Err.Number
    file.WriteLine "Error Description: " & Err.Description
    file.WriteLine "Source: " & Err.Source
    file.Close
    Set file = Nothing
    Set fso = Nothing
    'Return a default value or handle the error appropriately
    MyFunction = -1 ' Indicate an error
  Else
    'Code to execute if no error occurred
  End If
  On Error GoTo 0 'Disable error handling
End Function