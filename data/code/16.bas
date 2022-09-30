Attribute VB_Name = "Helper16"
Option Explicit

Function GetMonthName(ByVal monthNum&) As String
    ' Gets month name by it number in year
    GetMonthName = Format(DateSerial(CStr(Year(Date)), monthNum, 1), "mmmm")
End Function
