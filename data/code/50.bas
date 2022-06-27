Attribute VB_Name = "Helper50"
Option Explicit

Function GetMaxValue(ByVal x As Variant, ByVal y As Variant) As Variant
    ' ф-ция возвращает максимальное значение из 2-х
    GetMaxValue = IIf(x > y, x, y)
End Function