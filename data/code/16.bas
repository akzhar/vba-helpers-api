Attribute VB_Name = "VbaHelper_GetMonthName"
Option Explicit

Function GetMonthName(ByVal monthNum&) As String
    ' Gets month name by it number in year
    GetMonthName = Format(DateSerial(CStr(Year(Date)), monthNum, 1), "mmmm")
End Function
