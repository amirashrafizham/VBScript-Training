On Error Resume Next
Err.Raise(7)
Set objFsoObject = WScript.CreateObject("Scripting.FileSystemObject")
If (objFsoObject.FileExists("C:\Users\amirashraf.ahmadizh\OneDrive - PETRONAS\Desktop\VBScript Training\Chapter 9\ScriptLog.txt")) Then
Set objLogFile = objFsoObject.OpenTextFile("C:\Users\amirashraf.ahmadizh\OneDrive - PETRONAS\Desktop\VBScript Training\Chapter 9\ScriptLog.txt", 8)
Else
Set objLogFile = objFsoObject.OpenTextFile("C:\Users\amirashraf.ahmadizh\OneDrive - PETRONAS\Desktop\VBScript Training\Chapter 9\ScriptLog.txt", 2, "True")
End If
objLogFile.WriteLine ("Date: "&Date&", Time: "&Time&vbCrLf&"Test.vbs Error: " & Err.Number & ", Description = " & _
Err.Description & " , Source = " & Err.Source)
objLogFile.Close()