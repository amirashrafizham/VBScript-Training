Option Explicit
Dim FsoObject, DiskDrive, AvailSpace

Set FsoObject = WScript.CreateObject("Scripting.FileSystemObject")
Set DiskDrive = FsoObject.GetDrive(FsoObject.GetDriveName("c:"))

AvailSpace = (DiskDrive.FreeSpace / 1024) / 1024
AvailSpace = FormatNumber(AvailSpace,0)
WScript.Echo ("You need 100 MB of free space to play this game."&_
vbCrLf &"Total amount of free space is currently: "& AvailSpace & " MB")