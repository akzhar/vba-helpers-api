Attribute VB_Name = "Helper18"
Option Explicit

Function GetWeekdayName(ByVal d$) As String
    ' ф-ция возвращает название дня недели по дате
    GetWeekdayName = weekDayName(weekDay(CDate(d), vbMonday))
End Function