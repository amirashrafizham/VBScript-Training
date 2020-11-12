Option Explicit

Dim input
Dim computer
Dim strResults

input = InputBox("Type rock, paper or scissors","Rock, Paper, Scissors")
'input = "rock"
Randomize
computer = Int((rnd*3)+1)

If computer = 1 Then
    computer = "rock"
    ElseIf computer = 2 Then
    computer = "paper"
    ElseIf computer = 3 Then
    computer = "scissors"
End If
WScript.Echo("You have picked "&input)
WScript.Echo("The computer has picked "&computer)   

Select Case input
Case "rock"
If computer= "rock" Then WScript.Echo "Tie"
If computer= "paper" Then WScript.Echo "You Lose"
If computer= "scissors" Then WScript.Echo "You Win"
Case "scissors"
If computer= "rock" Then WScript.Echo "You Lose"
If computer= "paper" Then WScript.Echo "You Win"
If computer= "scissors" Then WScript.Echo "Tie"
Case "paper"
If computer= "rock" Then WScript.Echo "You Win"
If computer= "paper" Then WScript.Echo "You Lose"
If computer= "scissors" Then WScript.Echo "Tie"
Case Else
WScript.Echo "Please input a correct statement"
WScript.Quit
End Select

'Remember if you use Select Case, your If and Then must be in the same line or else the programme doesn't work