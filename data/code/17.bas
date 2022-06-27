Attribute VB_Name = "Helper17"
Option Explicit

Function GetMonthByName(ByVal monthName$) As Long
    ' ф-ция возвращает порядковый номер месяца в году по его имени
    Dim monthNames(): monthNames = Array("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")
    GetMonthByName = GetIndexOf(monthNames, monthName) + 1 ' @(id 11)
End Function