Option Explicit
'Defining the variables
Dim intGuess
Dim intCounter
Dim intRandom
Dim intMax
Dim intMin
Dim strTitle 

strTitle = "Number guessing game"

'Defining the boundaries of random number generator
intMax = 5
intMin = 1
Randomize
intRandom = Int((5*Rnd) + 1) 



'Continously capture guess number while it is not equals to random number
Do 
' Capture input from the user 
    Do
    intGuess = InputBox("Type a number between "&intMin&" and "&intMax&" :",strTitle)
        If (isNumeric(intGuess) = False) Then
        MsgBox("Please input a number"),vbOkOnly,strTitle
        End If
    Loop While(isNumeric(intGuess) = False)

' Convert into Integer
    intGuess = CInt(intGuess)
    Wscript.Echo "The guess number is "&intGuess
    Wscript.Echo "The actual number is "&intRandom

' Check whether Guess is equivalent to Random or not
    If intGuess = intRandom  Then 
    MsgBox ("Congrats you are correct"),vbOkOnly,strTitle
    Exit Do
    ElseIf intGuess < intRandom Then 
    MsgBox("Guess number is lower than actual number"),vbOkOnly,strTitle 
    ElseIf intGuess > intRandom Then 
    MsgBox("Guess number is higher than actual number"),vbOkOnly,strTitle
    End If

Loop While intGuess <> intRandom