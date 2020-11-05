Option Explicit

Dim intCounter 'variable used to control For...Each loop
Dim strMessage 'Message to be displayed in a pop-up dialog

strMessage = "The array contains the following default game information " &vbCrLf&vbCrLf

Dim arrGame(4)
arrGame(0) = "Amir Ashraf" 'Default username
arrGame(1) = "Manchester" 'A place worth visiting
arrGame(2) = "Raspberry PI" 'An interesting object
arrGame(3) = "Tommy Shelby" 'An idol
arrGame(4) = "Red Velvet" 'A favourite dessert

For Each intCounter in arrGame
    strMessage = strMessage & intCounter & vbCrLf
Next

WScript.Echo strMessage

