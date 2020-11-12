Option Explicit


'************** FIRST  EXAMPLE **************
Dim counter
Dim ending 
Dim food
Dim input
ending = 3
food = " "

'For counter = 0 To ending
'    input= InputBox("What is your favourite food")
'    food = food&""&input&", "  
'Next
'WScript.Echo ("You like to eat:"&food)

'************** SECOND  EXAMPLE **************
Dim input2

For counter = 0 to 3 
    input2 = InputBox("What is your favourite food")    
    If input2 = "fuck" Then
    Wscript.Echo "You have input the f word"
    Exit For
    End if 
Next