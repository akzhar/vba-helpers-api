Attribute VB_Name = "Helper65"
Option Explicit

Function GetLastColumn(ByRef ws As Worksheet, ByVal rowNum&) As Long
    ' Gets number of last filled (not empty) column in the specified row on the specified worksheet
    GetLastColumn = ws.Cells(rowNum, Columns.Count).End(xlToLeft).Column
End Function