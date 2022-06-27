Attribute VB_Name = "Helper64"
Option Explicit

Function GetLastRow(ByRef ws As Worksheet, ByVal colNo&) As Long
    ' ф-ция возвращает номер последней не пустой строки по номеру колонки
    GetLastRow = ws.Cells(Rows.Count, colNo).End(xlUp).row
End Function