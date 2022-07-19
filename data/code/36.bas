Attribute VB_Name = "Helper36"
Option Explicit

Function IsColoredLike(ByRef rng as Range, ByVal color) as Boolean
     ' ф-ция проверяет покрашен ли rng в указанный цвет
     Dim isHex as Boolean: isHex = Includes(CStr(color), "#") ' @(id 69)
     IsColoredLike = CBool(rng.Interior.Color = Iif(isHex, Hex2Long(color), color)) ' @(id 38)
End Function
