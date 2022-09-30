Attribute VB_Name = "Helper32"
Option Explicit

Function SetFontColor(ByRef rng As Range, ByVal color)
    ' Sets text color for specified range
    Dim isHex as Boolean: isHex = Includes(CStr(color), "#") ' @(id 69)
    rng.Font.Color = Iif(isHex, Hex2Long(color), color) ' @(id 38)
End Function