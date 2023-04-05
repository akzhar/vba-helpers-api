Attribute VB_Name = "Helper32"
Option Explicit

Function SetFontColor(ByRef rng As Range, ByVal color)
    ' Sets text color for specified range
    Dim isHex as Boolean: isHex = Includes(CStr(color), "#") ' @dependency: 69.bas
    rng.Font.Color = Iif(isHex, Hex2Long(color), color) ' @dependency: 38.bas
End Function