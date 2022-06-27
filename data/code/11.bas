Attribute VB_Name = "Helper11"
Option Explicit

Function GetIndexOf(ByRef arr(), ByVal element, Optional ByVal elementPos&) As Long
    ' ф-ция возвращает индекс элемента в массиве arr или -1, если эл-та в массиве нет
    Dim i&, arrElement As Variant
    For i = LBound(arr) To UBound(arr)
        If (IsNull(elementPos)) Then
            arrElement = arr(i)
        Else
            arrElement = arr(i, elementPos)
        End If
        If (element = arrElement) Then
            GetIndexOf = i
            Exit Function
        End If
    Next i
    GetIndexOf = - 1
End Function