Attribute VB_Name = "VbaHelper_GetLastRow"
Option Explicit

Function GetLastRow(ByRef ws As Worksheet, Optional ByVal colNum&) As Long
    ' Gets number of last filled (not empty) row in the specified column on the specified worksheet
    If colNum = 0 Then
        GetLastRow = ws.UsedRange.Rows.Count
    Else
        GetLastRow = ws.Cells(Rows.Count, colNum).End(xlUp).row
    End If
End Function