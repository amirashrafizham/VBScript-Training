'Single Dimension Arrays

Dim guessList(4)
guessList(0) = "Amir"
guessList(1) = "Ashraf"
guessList(2) = "Ahmad"
guessList(3) = "Izham"

WScript.Echo(guessList(1))

'Multi Dimensional Arrays
Dim arrguessList(3,1)
arrguessList(0,0) = "Amir"
arrguessList(0,1) = "0124206879"
arrguessList(1,0) = "Ashraf"
arrguessList(1,1) = "0121234567"
arrguessList(2,0) = "Ahmad"
arrguessList(2,1) = "0129876543"
arrguessList(3,0) = "Izham"
arrguessList(3,1) = "0126663639"

'*************************************************************************
' How (3,1) arrays look like

'|0,0|0,1|
'|1,0|1,1|
'|2,0|2,1|
'|3,0|3,1|
'*************************************************************************

