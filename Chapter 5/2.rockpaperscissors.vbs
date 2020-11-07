Option Explicit

Dim input
Dim computer

input = InputBox("Type rock, paper or scissors")
'input = "scissors"
Randomize
computer = Int((rnd*3)+1)
'WScript.Echo(computer)

If computer = 1 Then
    computer = "rock"
    ElseIf computer = 2 Then
    computer = "paper"
    ElseIf computer = 3 Then
    computer = "scissors"
End If
WScript.Echo(computer)   

If (input = "scissors") AND (computer = "rock") Then
    WScript.Echo ("Lose")
ElseIf (input = "scissors") AND (computer = "paper") Then
    WScript.Echo ("Win")
ElseIf (input = "scissors") AND (computer = "scissors") Then
    WScript.Echo ("Tie")
End If


If (input = "paper") AND (computer = "rock") Then
    WScript.Echo ("Win")
ElseIf (input = "paper") AND (computer = "paper") Then
    WScript.Echo ("Tie")
ElseIf (input = "paper") AND (computer = "scissors") Then
    WScript.Echo ("Lose")
End If


If (input = "rock") AND (computer = "rock") Then
    WScript.Echo ("Tie")
ElseIf (input = "rock") AND (computer = "paper") Then
    WScript.Echo ("Lose")
ElseIf (input = "rock") AND (computer = "scissors") Then
    WScript.Echo ("Win")
End If