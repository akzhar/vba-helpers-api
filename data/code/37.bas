Attribute VB_Name = "Helper37"
Option Explicit

Function Rgb2Long(ByVal R&, ByVal G&, ByVal B&) As Long
    ' ф-ция возвращает значение RGB цвета в Long формате
    Rgb2Long = (B * 65536) + (G * 256) + R
End Function