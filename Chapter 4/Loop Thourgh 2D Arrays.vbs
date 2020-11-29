Dim activities(4,1)
activities(0,0)="Work" 
activities(0,1)= 9
activities(1,0)="Eat"
activities(1,1)=1
activities(2,0)="Commute"
activities(2,1)=2
activities(3,0)="Play Game"
activities(3,1)=1
activities(4,0)="Sleep"
activities(4,1)=7

Dim intArrayDim1
Dim intArrayDim2
Dim intCounter1
Dim intCounter2

intArrayDim1 = UBound(activities,1)
intArrayDim2 = UBound(activities,2)

WScript.Echo intArrayDim1
WScript.Echo intArrayDim2

For intCounter1 = 0 To intArrayDim1
    For intCounter2 = 0 to intArrayDim2 
    WSCript.Echo ("["&intCounter1&","&intCounter2&"]"&activities(intCounter1,intCounter2))
    Next
Next

'Index    |0          | 1
'0        |"Work"     |	9
'1   	  |"Eat"      |	1
'2	      |"Commute"  | 2
'3	      |"Play Game"|	1
'4	      |"Sleep"	  | 7