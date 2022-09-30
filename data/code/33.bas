Attribute VB_Name = "Helper33"
Option Explicit

Function SetBackColor(ByRef rng As Range, ByVal color)
    ' Sets background color for specified range
    Dim isHex as Boolean: isHex = Includes(CStr(color), "#") ' @(id 69)
    rng.Interior.Color = Iif(isHex, Hex2Long(color), color) ' @(id 38)
End Function