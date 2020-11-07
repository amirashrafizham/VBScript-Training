Option Explicit

Dim pay

pay = 250

If pay = 250 Then
    If Weekday(date()) = 1 Then
        WScript.Echo "It's Sunday, the TV store is closed on Sundays."
    Else
    WScript.Echo "Go buy that TV!"
    End If
Else
    WScript.Echo "Save your money in the bank"
End If