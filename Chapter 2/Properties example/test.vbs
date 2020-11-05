Set WshNtwk = WScript.CreateObject("WScript.Network")

WScript.Echo "The User Domain is "&WshNtwk.UserDomain&_
vbCrLf&"The Computer Name is "&WshNtwk.ComputerName&_
vbCrLf&"The User Name is "&WshNtwk.UserName
