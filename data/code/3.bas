Attribute VB_Name = "Helper3"
Option Explicit

Function Rng2Array(ByRef rng As Range) As Variant()
    ' Converts range to array
    Dim cell As Range, arr()

    For Each cell In rng
        If cell.Value <> "" Then
            Call AddToArr(arr, CStr(cell.Value))  ' @(id 1)
        End If
    Next cell

    Rng2Array = arr
End Function