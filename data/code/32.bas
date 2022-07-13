Attribute VB_Name = "Helper32"
Option Explicit

Function SetRngFontColor(ByRef rng As Range, ByVal hexColor$)
    ' ф-ция устанавливает цвет текста у rng
    rng.Font.Color = Hex2Long(hexColor) ' @(id 38)
End Function