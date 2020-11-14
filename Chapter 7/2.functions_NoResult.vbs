Option Explicit
Dim strMessage
Dim strTitle
strMessage =    "Thank you for playing the game "&_
                vbCrLf&vbCrLf&"Please play again soon"
strTitle = "Test Game"

DisplaySplashScreen()

Function DisplaySplashScreen()
    MsgBox strMessage, 4144, strTitle 
End Function