Attribute VB_Name = "Helper2"
Option Explicit

Function GetArrLength(ByRef arr) As Long
    ' ф-ция возвращает длину 1 мерного массива arr
    If IsEmpty(arr) Then
        GetArrLength = 0
    Else
        GetArrLength = UBound(arr) - LBound(arr) + 1
    End If
End Function