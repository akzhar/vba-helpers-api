Attribute VB_Name = "Helper17"
Option Explicit

Function GetMonthNum(ByVal monthName$) As Long
    ' Gets month number in year by its name
    Dim monthNames(): monthNames = Array("january", "ferbuary", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december")
    GetMonthNum = GetIndexOf(monthNames, LCase(monthName)) + 1 ' @dependency: 11.bas
End Function