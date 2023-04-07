Attribute VB_Name = "VbaHelper_SetBackColor"
Option Explicit

Function SetBackColor(ByRef rng As Range, ByVal color)
    ' Sets background color for specified range
    Dim isHex as Boolean: isHex = HasSubstring(CStr(color), "#") ' @dependency: 69.bas
    rng.Interior.Color = Iif(isHex, Hex2Long(color), color) ' @dependency: 38.bas
End Function