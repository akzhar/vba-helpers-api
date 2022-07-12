Attribute VB_Name = "Helper55"
Option Explicit

Function TurnFiltersOn(ByRef ws As Worksheet, ByVal headerRow&)
    ' ф-ция активирует фильтры на листе
    Dim lastCol&: lastCol = GetLastColumn(ws, headerRow) ' @(id 65)
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).AutoFilter
End Function