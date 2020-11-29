On Error Resume Next
NonExistentFunction()

If Err > 0 then
    Err.Number = 9999
    Err.Description = "This script is still a work in progress."
    MsgBox "Error: "&Err.Number&" - "&Err.Description
    Err.Clear
End If

'Error has 2 properties i.e. Number and Description
'Error has 2 methods i.e. Err.Clear() and Err.Raise()
