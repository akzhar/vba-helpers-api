Attribute VB_Name = "Helper4"
Option Explicit

Function IsInArray(ByRef arr(), ByVal element) As Boolean
    ' ф-ция проверяет наличие element в 1 мерном массиве arr
    Dim i&
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = element Then
            IsInArray = True
            Exit Function
        End If
    Next i

    IsInArray = False
End Function
