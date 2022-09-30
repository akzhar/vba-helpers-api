Attribute VB_Name = "Helper66"
Option Explicit

Function GetEdgeRows(ByRef rng As Range) As Long()
    ' Gets the numbers of edge rows in the range (from ... to ...)
    Dim firstRow&: firstRow = rng.Rows(1).row
    Dim lastRow&: lastRow = rng.Rows.Count + firstRow - 1
    Dim arr(1) As Long
    arr(0) = firstRow
    arr(1) = lastRow
    GetEdgeRows = arr
End Function
