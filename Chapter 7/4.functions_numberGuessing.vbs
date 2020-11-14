Option Explicit
'Defining the variables
Public intGuess
Dim intCounter
Public intRandom
Dim intMax
Dim intMin
Public strTitle 
Public intEnding
strTitle = "Number guessing game"


RandomNumber()
PlayTheGame()
'***************************************************************************************************
'FUNCTIONS AND PROCEDURES 

Function RandomNumber()
    intMax = 5
    intMin = 1
    Randomize
    intRandom = Int((5*Rnd) + 1) 
End Function


Function PlaytheGame()
 Do
    'Capture input from the user 
        ValidateAnswer()         
    'Convert into Integer
        NumberConversion()
    'Test Answer
        TestAnswer()
    'Exit Game
        If intEnding = 1 Then
        Exit Do
        End If

 Loop While intGuess <> intRandom   
End Function

Function ValidateAnswer()
    Do
        intGuess = InputBox("Type a number between "&intMin&" and "&intMax&" :",strTitle)
        If (isNumeric(intGuess) = False) Then
        MsgBox("Please input a number"),vbOkOnly,strTitle
        End If
    Loop While(isNumeric(intGuess) = False)
End Function

Function NumberConversion()
    intGuess = CInt(intGuess)
    Wscript.Echo "The guess number is "&intGuess
    Wscript.Echo "The actual number is "&intRandom
End Function 
 
Function TestAnswer()
    ' Check whether Guess is equivalent to Random or not
    If intGuess = intRandom  Then 
    MsgBox ("Congrats you are correct"),vbOkOnly,strTitle
    intEnding = 1
    ElseIf intGuess < intRandom Then 
    MsgBox("Guess number is lower than actual number"),vbOkOnly,strTitle 
    intEnding = 0
    ElseIf intGuess > intRandom Then 
    MsgBox("Guess number is higher than actual number"),vbOkOnly,strTitle
    intEnding = 0
    End If
End Function