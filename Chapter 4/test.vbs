Option Explicit
'
'Const cTitleBarMsg = "The Three Little Pigs"
'
'WScript.Echo ("Unce upon a time...."&_'
'        vbCrLf&"There were three litte pigs"&_
' '       vbCrLf&"Who liked to build things")
'
'Dim TodaysDate       
'TodaysDate = Weekday(Date)
'WScript.Echo ("Today is day number "&TodaysDate)
'
'If TodaysDate = vbThursday Then
'    WScript.Echo ("Today is Thursday")
'Else
'    WScript.Echo ("Today is not Thursday")
'End If

'Dim intMyAge
'intMyAge = 37
'WScript.Echo "I am "&intMyAge
'intMyAge = intMyAge + 1
'WScript.Echo "Next year I will be "&intMyAge
'intMyAge = intMyAge - 1
'intMyAge = (intMyAge/5) + 3*10
'WScript.Echo "If I took that value, divided it by 5 and add 30: "&intMyAge

'Dim itsMyParty(9)
'itsMyParty(1) = "Amir"
'itsMyParty(2) = "Ashraf"


'Dim astrItsMyParty(3,1)
'astrItsMyParty(0,0) = "Jerry"
'astrItsMyParty(0,1) = "550-9933"
'astrItsMyParty(1,0) = "Molly"
'astrItsMyParty(1,1) = "550-8876"
'astrItsMyParty(2,0) = "William"
'astrItsMyParty(2,1) = "697-4443"
'astrItsMyParty(3,0) = "Alexander"
'astrItsMyParty(3,1) = "696-4344"

'WScript.Echo astrItsMyParty(3,0)
'WScript.Echo astrItsMyParty(1,1)


Dim intCounter 
Dim strMessage
Dim intSize
Dim strMessage2
strMessage = "The array contains the following default game " &_
             "information: " &vbCrLf &vbCrLf

strMessage2 = "The array contains the following default game " &_
              "information: " &vbCrLf &vbCrLf

Dim astrGameArray(4)
astrGameArray(0) = "Joe Blow"
astrGameArray(1) = "Nevada"
astrGameArray(2) = "Soda Can"
astrGameArray(3) = "Barney"
astrGameArray(4) = "Pickle"

For Each intCounter in astrGameArray
    strMessage = strMessage & intCounter & vbCrLf
Next

WScript.Echo strMessage

intSize = UBound(astrGameArray)
WScript.Echo (intSize)

For intCounter = 0 To intSize
    strMessage2 = strMessage2 &astrGameArray(intCounter)&vbCrLf 
Next

WScript.Echo (strMessage2)