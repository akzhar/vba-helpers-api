Attribute VB_Name = "Helper50"
Option Explicit

Function GetMax(ByVal x As Variant, ByVal y As Variant) As Variant
    ' ф-ция возвращает максимальное значение из 2-х
    GetMax = IIf(x > y, x, y)
End Function