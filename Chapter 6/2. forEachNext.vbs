Option Explicit

Dim objFsoObject, objFolderName
Dim strMember, strFileList

Set objFsoObject = CreateObject("Scripting.FileSystemObject")
Set objFolderName = objFsoObject.GetFolder("C:\Users\amirashraf.ahmadizh\OneDrive - PETRONAS\Desktop\VBScript Training\Chapter 4 ") 
For Each strMember in objFolderName.Files
    strFileList = strFileList & strMember.Name & vbCrLf
Next

WScript.Echo "These are the files in the VBScript Training folder"&vbCrLf&strFileList
