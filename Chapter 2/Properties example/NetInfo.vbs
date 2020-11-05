Set WshNtwk = WScript.CreateObject("WScript.Network")

userDomain = "User Domain" & vbTab & "=" & WshNtwk.UserDomain &_
            vbCrLf & "Computer Name" & vbTab & "=" & WshNtwk.ComputerName &_
            vbCrLf & "User Name" & vbTab & "= " & WshNtwk.UserName

'computerName ="Computer Name" & vbTab & "=" & WshNtwk.ComputerName
'userName = "User Name" & vbTab & "= " & WshNtwk.UserName

MsgBox (userDomain)


