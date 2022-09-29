Attribute VB_Name = "Helper91"
Option Explicit

Function GetColumnNumber(ByVal letter$) As Long
  ' ф-ция возвращает номер столбца по его букве
  GetColumnNumber = Range(letter & 1).Column
End Function