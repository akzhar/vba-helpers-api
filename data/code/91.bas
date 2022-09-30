Attribute VB_Name = "Helper91"
Option Explicit

Function GetColumnNumber(ByVal letter$) As Long
  ' Gets column's number by its letter
  GetColumnNumber = Range(letter & 1).Column
End Function