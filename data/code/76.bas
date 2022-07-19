Attribute VB_Name = "Helper76"
Option Explicit

Function CONCATIF(ByRef rngToCheck As Range, ByRef rngToConcat As Range, ByVal pattern$, Optional separator$ = " ") As String
    ' аналог встроенной ф-ции CONCATENATE, но с возможностью задать условие конкатенации
    Application.Volatile True
    Dim cell As Range, str$
    For Each cell In rngToCheck
       If cell.Value Like pattern And Trim(rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)) <> "" Then
          str = str & IIf(str <> "", separator, "") & rngToConcat.Cells(cell.Row - rngToCheck.Row + 1, 1)
       End If
    Next cell
    CONCATIF = str
End Function