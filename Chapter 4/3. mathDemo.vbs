Option Explicit

Dim intMyAge
intMyAge = 26
WScript.Echo "I am " &intMyAge& " years old"

intMyAge = intMyAge + 1
WScript.Echo "Next year I will be " &intMyAge& " years old"

intMyAge = ((intMyAge/5) + 3) * 10
WScript.Echo    "And this is my age would be if I divide it by 5 add 3 and times 10, I would be "&_
                intMyAge&" years old"