MsgBox ("Welcome to Rock, Paper and Scissors game. Here are the rules of the game:"&_
vbCrLf&""&_
vbCrLf&"1. Guess the same things as the computer to tie."&_
vbCrLf&"2. Paper covers rock and wins."&_
vbCrLf&"3. Rock breaks scissors and wins"&_
vbCrLf&"4. Scissors cut paper and wins")

Dim computerInput
If random = 1 Then 
    computerInput = "paper"
ElseIf random = 2 Then 
computerInput = "rock"
Else computerInput = "scissors" 
End If
WScript.Echo (random&" "&computerInput) 
Dim userinput  
userinput = InputBox("Type paper, rock or scissors")

If userinput = computerInput Then   
MsgBox ("Congrats, you have won, the computer have picked "&computerInput)
Else
MsgBox ("Try again, the computer have picked "&computerInput)
End If 

