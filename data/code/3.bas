Attribute VB_Name = "VbaHelper_Rng2Array"
Option Explicit

Function Rng2Array(ByRef rng As Range) As Variant()
    ' Converts range to array
    Dim cell As Range, arr()

    For Each cell In rng.Cells
        If cell.Value <> "" Then
            Call AddToArr(arr, CStr(cell.Value))  ' @dependency: 1.bas
        End If
    Next cell

    Rng2Array = arr
End Function