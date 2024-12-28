Function GetObjectRobust(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    ' Attempt to create the object if not found
    On Error GoTo CreateError ' Handle potential errors during object creation
    Set obj = CreateObject(progID)
    On Error Resume Next  ' Reset error handling
  End If
  If obj Is Nothing Then
      WScript.Echo "Could not get or create object: " & progID
  End If
  Set GetObjectRobust = obj
  Exit Function
CreateError:
  WScript.Echo "Error creating object: " & Err.Description
  Err.Clear
  Set GetObjectRobust = Nothing
End Function

' Example usage:
Set excelApp = GetObjectRobust("Excel.Application")
if excelApp is nothing then
  WScript.Echo "Excel could not be accessed"
else
  WScript.Echo "Excel is running"
  excelApp.Quit
end if