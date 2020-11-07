'*************************************************************************
'Script Name: StarTrekQuiz.vbs
'Author: Amir Ashraf 
'Created: 7/11/20
'Description: This script creates a Star Trek Quiz game.
'*************************************************************************
Option Explicit
'Creating a splash screen 

Dim strSplashImage
Dim cTitleBarMsg
Dim intPlayGame
Dim intAnswer1
Dim intAnswer2
Dim intAnswer3
Dim intAnswer4
Dim intAnswer5
Dim intCounter
Dim intTotal
intCounter = 0
cTitleBarMsg = "Math quiz" 

strSplashImage = space(11) & "********" & vbCrLf & _
"******************" & space(20) & "**************************" & _
space(20) & vbCrLf & "*"& space(35) & "*" & space(18) & _
"**" & space(46) & "*" & vbCrLf & " ******************" & _
space(20) & "*************************" & vbCrLf & space(31) & _
"******" & space(26) & "***" & vbCrLf & _
space(34) & "******" & space(22) & "***" & vbCrLf & _
space(37) & "******" & space(17) & "***" & vbCrLf & _
space(26) & " ****************************" & vbCrLf & _
space(26) & "*******************************" & vbCrLf & _
space(26) & "******************************" & vbCrLf & _
space(26) & " ****************************" & vbCrLf & vbCrLf & vbCrLf &_
space(10) & "Would you like to test your math skills?"


intPlayGame= MsgBox (strSplashImage,36,cTitleBarMsg)

If intPlayGame = 7 Then 
    MsgBox("Thank you and have a good day"),vbOkOnly,cTitleBarMsg
    WScript.Quit()
Else
    intAnswer1 = CInt(inputBox("What is 1 + 2?"))
    intAnswer2 = CInt(inputBox("What is 3 + 4?"))
    intAnswer3 = CInt(inputBox("What is 4 + 5?"))

    If intAnswer1 = 3 Then intCounter = intCounter + 1
    If intAnswer2 = 7 Then intCounter = intCounter + 1
    If intAnswer3 = 9 Then intCounter = intCounter + 1
End If
WScript.Echo (intCounter)
intTotal = Round(intCounter/3 * 100,2) 
MsgBox "Your total score is "&(intTotal),vbOkOnly,cTitleBarMsg&"%"