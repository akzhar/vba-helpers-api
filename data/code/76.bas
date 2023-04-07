Attribute VB_Name = "VbaHelper_CONCATIF"
Option Explicit

Function CONCATIF(ByRef rngToCheck As Range, ByRef rngToConcat As Range, ByVal pattern$, Optional separator$ = " ") As String
    ' Performs concatenation of values in a range by the condition
    Application.Volatile True
    Dim cell As Range, str$
    For Each cell In rngToCheck
       If cell.Value Like pattern And Trim(rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)) <> "" Then
          str = str & IIf(str <> "", separator, "") & rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)
       End If
    Next cell
    CONCATIF = str
End Function