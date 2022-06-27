Attribute VB_Name = "Helper36"
Option Explicit

Function CheckRngBackColor(ByRef rng as Range, ByVal hexColor$) as Boolean
     ' ф-ция проверяет покрашен ли rng в указанный цвет
     CheckRngBackColor = CBool(rng.Interior.Color = Hex2Rgb(hexColor)) ' @(id 38)
End Function
