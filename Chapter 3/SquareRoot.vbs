Dim number

Set WshSh1 = WScript.CreateObject("WScript.Shell")
WshSh1.Run("Excel")


number = InputBox("Please enter a number:", "Square Root Calculator")
MsgBox "The Square root of "&number&" is "&sqr(number),0,"Square Root Calculator"