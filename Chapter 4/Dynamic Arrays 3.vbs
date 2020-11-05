Option Explicit

Dim array ()
Dim firstArray
Dim secondArray
Dim intCounter
Dim combinedArray
Dim newCounter
newCounter = 0

firstArray = CInt(inputBox("Select the size of your first array")) 
ReDim array(firstArray)

For intCounter=0 To firstArray
    array(intCounter) = intCounter
    WScript.Echo (array(intCounter))
Next

secondArray = CInt(inputBox("Select the number of you array you would like to add on top of first array"))
combinedArray = firstArray + secondArray

WScript.Echo "Computing the combined array"
Redim Preserve array(combinedArray)
For intCounter = (firstArray+1) To combinedArray
    array(intCounter) = newCounter
    newCounter = newCounter + 1 
Next

For intCounter = 0 To combinedArray
    WScript.Echo (array(intCounter))
Next