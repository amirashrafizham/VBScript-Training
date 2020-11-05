Option Explicit

ReDim userArray(3)
dim arraySize
dim intCounter
dim strMessage

strMessage = "Below is the author's profile :"

userArray(0) = "Name : Amir Ashraf"
userArray(1) = "Alma Mater : Manchester (Msc) and Nottingham (BEng) "
userArray(2) = "Major : Mechanical Engineering"
userArray(3) = "Current Role: Software Engineering"

arraySize = UBound(userArray)

ReDim Preserve userArray(5) 
userArray(4) = "Interests : Jogging, Gym, Guitar, Reading Books, Digital Painting, Software Engineering, Embedded Electronics"
userArray(5) = "Goals: Software Engineering specializing in IoT and ML"

arraySize = UBound(userArray)

For intCounter = 0 To arraySize
    strMessage = strMessage&vbCrLf&vbCrLf&userArray(intCounter)
Next

MsgBox strMessage