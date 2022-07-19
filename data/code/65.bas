Attribute VB_Name = "Helper65"
Option Explicit

Function GetLastColumn(ByRef ws As Worksheet, ByVal rowNum&) As Long
    ' ф-ция возвращает номер последней не пустой колонки по номеру строки
    GetLastColumn = ws.Cells(rowNum, Columns.Count).End(xlToLeft).Column
End Function