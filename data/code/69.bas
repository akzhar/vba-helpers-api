Attribute VB_Name = "VbaHelper_HasSubstring"
Option Explicit

Function HasSubstring(ByVal str$, ByVal subStr$) As Boolean
    ' Checks if string includes substring
    HasSubstring = CBool(InStr(1, str, subStr, vbTextCompare) <> 0)
End Function