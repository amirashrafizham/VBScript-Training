Dim guessList()
ReDim guessList(0)

Dim intCounter, guessName
intCounter = 0
Do While UCase(guessName) <> "DONE"
    guessName = InputBox("Enter the name of someone to be invited: ")
    If UCase(guessName) <> "DONE" Then
    guessList(intCounter) = guessName
    Else
        Exit Do
    End If
intCounter = intCounter + 1
ReDim Preserve guessList(intCounter)
Loop 

Dim arrayCounter
Dim arraySize
arraySize = UBound(guessList)
Dim strMessage 
strMessage = "Below is the list of guess " &vbCrLf

For arrayCounter = 0 To arraySize
    strMessage = strMessage & guessList(arrayCounter) & vbCrLf
Next

WScript.Echo strMessage