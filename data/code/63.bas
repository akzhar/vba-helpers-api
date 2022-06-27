Attribute VB_Name = "Helper63"
Option Explicit

Function GetColumnLeterByNumber(ByVal colNo&) As String
  ' ф-ция возвращает букву столбца по его номеру
  GetColumnLeterByNumber = Replace(Cells(1, colNo).Address(0, 0), 1, "")
End Function