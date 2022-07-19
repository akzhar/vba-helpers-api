Attribute VB_Name = "Helper16"
Option Explicit

Function GetMonthName(ByVal monthNum&) As String
    ' ф-ция возвращает название месяца в формате "mmmm" по номеру месяца в году
    GetMonthName = Format(DateSerial(CStr(Year(Date)), monthNum, 1), "mmmm")
End Function
