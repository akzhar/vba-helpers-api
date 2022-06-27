Attribute VB_Name = "Helper81"
Option Explicit

Function Collection2Array(ByRef col As Object) As String()
    ' ф-ция конвертирует коллекцию в массив

    Dim i&, arr() As String

    For i = 0 To col.count - 1
        If i = 0 Then
            ReDim arr(0)
        Else
            ReDim Preserve arr(UBound(arr) + 1)
        End If
        arr(UBound(arr)) = col(i)
    Next i

    Collection2Array = arr
End Function