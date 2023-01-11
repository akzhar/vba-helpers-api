Attribute VB_Name = "Helper102"
Option Explicit

Function GetColumnNumber(ByVal colName$) As Long
    ' Gets column's number by its name
    GetColumnNumber = Range(colName & 1).Column
End Function