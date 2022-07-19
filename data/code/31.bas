Attribute VB_Name = "Helper31"
Option Explicit

Function IsColored(ByRef rng As Range) As Boolean
    ' ф-ция проверят окрашен ли диапазон (цвет заливки и цвет текста <> дефолт)
    IsColored = CBool(rng.Interior.ColorIndex <> xlColorIndexNone Or rng.Font.Color <> 0)
End Function
