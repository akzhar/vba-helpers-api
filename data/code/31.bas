Attribute VB_Name = "Helper31"
Option Explicit

Function IsColored(ByRef rng As Range) As Boolean
    ' Checks if range is colored
    IsColored = CBool(rng.Interior.ColorIndex <> xlColorIndexNone Or rng.Font.Color <> 0)
End Function
