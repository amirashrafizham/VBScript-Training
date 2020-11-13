Option Explicit

Dim intMissedGuesses, strPlayerAnswer
Dim strTitle
strTitle = "Peter Pan Game"
intMissedGuesses = 1

Do 
    strPlayerAnswer = inputBox("Where does Peter Pan live?",strTitle)

    If strPlayerAnswer = "neverland" Then
        MsgBox("You are correct !"),vbOkOnly,strTitle
        Exit Do
    Else
        MsgBox("Try again, you have "&(4-intMissedGuesses)&" tries left"),vbOkOnly,strTitle  
        intMissedGuesses = intMissedGuesses + 1
        If intMissedGuesses = 4 Then
        MsgBox("You have run out of tries"),vbOkOnly,strTitle
        End If 
    End If

Loop Until (intMissedGuesses > 3)