Attribute VB_Name = "Helper55"
Option Explicit

Function TurnFiltersOn(ByRef ws As Worksheet)
    ' ф-ция активирует фильтры на листе
    Dim lastCol&: lastCol = GetLastColumn(ws, 1) ' @(id 65)
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
End Function