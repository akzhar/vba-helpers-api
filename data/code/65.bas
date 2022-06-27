Attribute VB_Name = "Helper65"
Option Explicit

Function GetLastColumn(ByRef ws As Worksheet, ByVal rowNo&) As Long
    ' ф-ция возвращает номер последней не пустой колонки по номеру строки
    GetLastColumn = ws.Cells(rowNo, Columns.Count).End(xlToLeft).Column
End Function