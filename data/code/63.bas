Attribute VB_Name = "Helper63"
Option Explicit

Function GetColumnName(ByVal colNum&) As String
  ' Gets column's name by its number
  GetColumnName = Replace(Cells(1, colNum).Address(0, 0), 1, "")
End Function