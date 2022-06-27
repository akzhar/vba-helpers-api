Attribute VB_Name = "Helper33"
Option Explicit

Function SetRngBackColor(ByRef rng As Range, ByVal hexColor$)
    ' ф-ция устанавливает цвет заливки у rng
    rng.Interior.Color = Hex2Rgb(hexColor) ' @(id 38)
End Function