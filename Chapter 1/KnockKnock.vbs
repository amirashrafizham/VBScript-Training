Dim Reply1
Dim Reply2
Reply1 = InputBox("Knock Knock!")

If Reply1 = "Who's there?" Then
    Reply2 = InputBox("Panther!")
    If Reply2 = "Panther who?" Then
    MsgBox("Panther Pantat")
    ElseIf Reply2 <> "Panther who?" Then
    MsgBox("Try again reply2")    
    End If
ElseIf Reply1 <> "Who's there?" Then
MsgBox("Try again reply1")
End If



