Attribute VB_Name = "Helper55"
Option Explicit

Function TurnFiltersOn(ByRef ws As Worksheet, ByVal headerRow&)
    ' Turns on autofilters on the specified worksheet in the specified headers row
    Dim lastCol&: lastCol = GetLastColumn(ws, headerRow) ' @dependency: 65.bas
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, lastCol)).AutoFilter
End Function