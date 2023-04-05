Attribute VB_Name = "Helper36"
Option Explicit

Function IsColoredLike(ByRef rng as Range, ByVal color) as Boolean
     ' Checks if range's background is colored in specified color
     Dim isHex as Boolean: isHex = Includes(CStr(color), "#") ' @dependency: 69.bas
     IsColoredLike = CBool(rng.Interior.Color = Iif(isHex, Hex2Long(color), color)) ' @dependency: 38.bas
End Function
