On Error Resume Next
Err.Raise(7)
Set objWshShl = WScript.CreateObject("WScript.Shell")
objWshShl.LogEvent 1, "Date: "&Date&", Time: "&Time&vbCrLf&"Test.vbs Error: " & Err.Number & ", Description = " & _
Err.Description & " , Source = " & Err.Source