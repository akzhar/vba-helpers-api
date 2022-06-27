Attribute VB_Name = "Helper73"
Option Explicit

Function Rng2String(ByRef rng As Range) As String
    ' ф-ция склеивает диапазон в текстовую строку через разделитель
    Rng2String = Join(Application.Transpose(rng), Chr(10))
End Function