Option Explicit
'------------First Example of Do While Loop------------
Dim intCounter
intCounter = 0 

Do  
   WScript.Echo intCounter 
   intCounter = intCounter + 1 
 Loop While (intCounter < 10) 

'------------Second Example of Do While Loop------------
Dim input, list
Dim strTitleMsg
list = ""
strTitleMsg = "Name collection"

Do 
   input = inputBox("Please enter a name",strTitleMsg)
   list = list&input&vbCrLf 
   intCounter = intCounter + 1
Loop While (intCounter < 3)
MsgBox(list),vbOkOnly,strTitleMsg

'------------Third Example of Do While Loop------------
Do 
    input = inputBox("Please enter a name",strTitleMsg)
    list = list&input&vbCrLf 
    intCounter = intCounter + 1
    If intCounter = 3 Then 
        Exit Do
    End If 
Loop While (intCounter > -1)
MsgBox(list),vbOkOnly,strTitleMsg