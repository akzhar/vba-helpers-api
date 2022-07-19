Attribute VB_Name = "Helper18"
Option Explicit

Function GetWeekday(ByVal d As Date) As String
    ' ф-ция возвращает название дня недели по дате
    GetWeekday = WeekdayName(Weekday(d, vbMonday))
End Function