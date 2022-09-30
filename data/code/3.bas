Attribute VB_Name = "Helper3"
Option Explicit

Function Rng2Array(ByRef rng As Range) As String()
    ' Converts range to array
    Dim i&, cell As Range, arr() As String
    
    ReDim arr(rng.Count - 1)

    For Each cell In rng
        arr(i) = CStr(cell.Value)
        i = i + 1
    Next cell
    
    Rng2Array = arr
End Function