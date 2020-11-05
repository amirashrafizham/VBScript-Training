Option Explicit

Dim objWshObject

Set objWshObject = CreateObject("WScript.Shell")

WScript.Echo "This computer is running on "&objWshObject.ExpandEnvironmentStrings("%OS%")&_
    " and has "&objWshObject.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")& " processors"