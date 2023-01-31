Attribute VB_Name = "Helper65"
Option Explicit

Function GetLastColumn(ByRef ws As Worksheet, Optional ByVal rowNum&) As Long
    ' Gets number of last filled (not empty) column in the specified row on the specified worksheet
    If rowNum = 0 Then
        GetLastColumn = ws.UsedRange.Columns.Count
    Else
        GetLastColumn = ws.Cells(rowNum, Columns.Count).End(xlToLeft).Column
    End If
End Function