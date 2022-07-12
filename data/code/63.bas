Attribute VB_Name = "Helper63"
Option Explicit

Function GetColumnLeterByNum(ByVal colNo&) As String
  ' ф-ция возвращает букву столбца по его номеру
  GetColumnLeterByNum = Replace(Cells(1, colNo).Address(0, 0), 1, "")
End Function