Function GetObject(progID)
  On Error Resume Next
  Set obj = GetObject(progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = CreateObject(progID)
  End If
  Set GetObject = obj
End Function

'Example usage
Set myExcel = GetObject("Excel.Application")
if myExcel is nothing then
  WScript.Echo "Excel is not running"
else
  WScript.Echo "Excel is running"
  myExcel.Quit
end if