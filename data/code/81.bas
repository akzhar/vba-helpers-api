Attribute VB_Name = "Helper81"
Option Explicit

Function Collection2Array(ByRef coll As Object) As Variant()
    ' ф-ция конвертирует коллекцию в массив

    Dim i&, arr()

    For i = 1 To coll.Count
        If i = 1 Then
            ReDim arr(0)
        Else
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        arr(UBound(arr)) = coll(i)
    Next i

    Collection2Array = arr
End Function