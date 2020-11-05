Option Explicit

Dim arrTestArray() 
Dim firstInput
Dim secondInput
Dim intCounter

firstInput= inputBox ("Please select the initial size of your 1st array")
WScript.Echo ("The user has picked an array of size "&firstInput)
ReDim arrTestArray(firstInput)

For intCounter = 0 To firstInput   
    arrTestArray(intCounter) = intCounter
    WScript.Echo(arrTestArray(intCounter))
Next

secondInput = inputBox("Please select a number to resize of your array again")
ReDim Preserve arrTestArray(secondInput)
WScript.Echo ("The user has re-picked an array of size "&secondInput)

Dim newCounter 
newCounter = secondInput - firstInput
For intCounter = newCounter To secondInput 
    arrTestArray(intCounter) = intCounter 
Next


For intCounter = 0 To secondInput
    WScript.Echo (arrTestArray(intCounter))
Next