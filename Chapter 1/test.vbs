Greetings = InputBox("What is your name?")
MsgBox ("Hello "&Greetings&". Welcome to the Knock Knock joke game")
reply1 = InputBox("Knock Knock")
If reply1 = "whos there" Then
    reply2 = InputBox("Panther")
    If reply2 = "panther who" Then
    MsgBox ("Panther no panths I'm going swimming")
    Else
    MsgBox ("Try again")
    End If
Else 
MsgBox ("Try again")
End If
