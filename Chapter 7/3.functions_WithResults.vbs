Option Explicit
Dim strPlayersName 
Dim strGetPlayersName
strPlayersName = GetPlayersName()

WScript.Echo ("The player's name is "&strPlayersName)

Function GetPlayersName()
    GetPlayersName = InputBox("What is your first name")
End Function