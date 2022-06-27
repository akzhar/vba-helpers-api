Attribute VB_Name = "Helper1"
Option Explicit

Function AddToArr(ByRef arr(), ByVal element)
    ' ф-ция добавляет element в 1 мерный массив arr
    If (Not arr) = -1 Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    arr(UBound(arr)) = element
End Function