Option Explicit

dim userArray(3)
dim arraySize
dim intCounter
dim strMessage

strMessage = "Below is the author's profile :"

userArray(0) = "Name : Amir Ashraf"
userArray(1) = "University : Manchester"
userArray(2) = "Major : Mechanical Engineering"
userArray(3) = "Current Role: Software Engineering"

arraySize = UBound(userArray)

For intCounter = 0 To arraySize
    strMessage = strMessage&vbCrLf&vbCrLf&userArray(intCounter)
Next

MsgBox strMessage