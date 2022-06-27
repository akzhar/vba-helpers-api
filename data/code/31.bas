Attribute VB_Name = "Helper31"
Option Explicit

Function CheckIfRngColored(ByRef rng As Range) As Boolean
    ' ф-ция проверят окрашен ли диапазон (цвет заливки и цвет текста <> дефолт)
    CheckIfRngColored = CBool(rng.Interior.ColorIndex <> xlNone Or rng.Font.ColorIndex <> xlAutomatic)
End Function
