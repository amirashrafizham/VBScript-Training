On Error Resume Next
WScrp.Echo " Hello world!"
On Error Goto 0 
WScrip.Echo " Goodbye world!"


'On Error Goto 0 nullifies the effects of the On Error Resume Next statements. 
'The first error is ignored, but the second error is reported, hence causing the script to halt