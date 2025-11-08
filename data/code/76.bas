Attribute VB_Name = "VbaHelper_CONCATIF"
Option Explicit

Function CONCATIF(ByRef rngToCheck As Range, ByRef rngToConcat As Range, ByVal pattern$, Optional separator$ = " ") As String
    ' Performs concatenation of values in a range by the condition
    Application.Volatile True
    Dim rng As Range, str$, i&
    For i = 1 To rngToCheck.Cells.Count
        Set rng = rngToCheck(i)
        If rng.Value Like pattern And Trim(rngToConcat(i).Value) <> "" Then
            str = str & IIf(str <> "", separator, "") & rngToConcat(i).Value
        End If
    Next i
    CONCATIF = str
End Function