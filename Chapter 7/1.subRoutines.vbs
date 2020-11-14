Option Explicit
Dim strMessage
strMessage = "Thank you for playing the game"&_
             vbCrLf&vbCrLf&"Please play again soon!"

DsiplaySplashScreen()

Sub DsiplaySplashScreen()
    MsgBox strMessage,4144,"Test Game"
End Sub