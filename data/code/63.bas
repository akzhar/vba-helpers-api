Attribute VB_Name = "Helper63"
Option Explicit

Function GetColumnLeter(ByVal colNum&) As String
  ' ф-ция возвращает букву столбца по его номеру
  GetColumnLeter = Replace(Cells(1, colNum).Address(0, 0), 1, "")
End Function