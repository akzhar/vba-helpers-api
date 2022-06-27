Attribute VB_Name = "Helper37"
Option Explicit

Function GetRgbLongValue(ByVal R&, ByVal G&, ByVal B&) As Long
    ' ф-ция возвращает значение RGB цвета в Long формате
    GetRgbLongValue = (R * 65536) + (G * 256) + B
End Function