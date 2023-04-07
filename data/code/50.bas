Attribute VB_Name = "VbaHelper_GetMax"
Option Explicit

Function GetMax(ByVal x As Variant, ByVal y As Variant) As Variant
    ' Gets max value from two specified values
    GetMax = IIf(x > y, x, y)
End Function