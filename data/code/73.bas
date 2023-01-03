Attribute VB_Name = "Helper73"
Option Explicit

Function Rng2String(ByRef rng As Range, ByVal delimiter$, Optional ByVal fnName$) As String
    ' Concatenates all values from the range into a string separated by the specified delimiter
    ' If fnName is provided - concatenates only cells that passed the function
    Rng2String = ""
    If fnName = "" Then
        Rng2String = Join(Rng2Array(rng), delimiter) ' @(id 3)
    Else
        Dim cell As Range
        For Each cell In rng
            If Application.Run(fnName, cell.value) Then
                Rng2String = Rng2String & delimiter & cell.value
            End If
        Next cell
        Rng2String = Right(Rng2String, Len(Rng2String) - 1) ' remove last delimiter
    End If
End Function