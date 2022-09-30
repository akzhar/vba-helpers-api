Attribute VB_Name = "Helper64"
Option Explicit

Function GetLastRow(ByRef ws As Worksheet, ByVal colNum&) As Long
    ' Gets number of last filled (not empty) row in the specified column on the specified worksheet
    GetLastRow = ws.Cells(Rows.Count, colNum).End(xlUp).row
End Function