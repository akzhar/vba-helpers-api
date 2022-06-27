Attribute VB_Name = "Helper66"
Option Explicit

Function GetRngRowsRange(ByRef rng As Range) As Long()
    ' ф-ция возвращает массив из двух значений: 0 - первая строка rng, 1 - последняя строка rng
    Dim firstRow&: firstRow = rng.Rows(1).row
    Dim lastRow&: lastRow = rng.Rows.Count + firstRow - 1
    Dim arr(1) As Long
    arr(0) = firstRow
    arr(1) = lastRow
    GetRngRowsRange= arr
End Function
