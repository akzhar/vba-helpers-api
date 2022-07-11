Attribute VB_Name = "Helper18"
Option Explicit

Function GetWeekdayName(ByVal d As Date) As String
    ' ф-ция возвращает название дня недели по дате
    GetWeekdayName = WeekdayName(Weekday(d, vbMonday))
End Function