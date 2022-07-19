Attribute VB_Name = "Helper17"
Option Explicit

Function GetMonthNum(ByVal monthName$) As Long
    ' ф-ция возвращает порядковый номер месяца в году по его имени
    Dim monthNames(): monthNames = Array("январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь")
    GetMonthNum = GetIndexOf(monthNames, LCase(monthName)) + 1 ' @(id 11)
End Function