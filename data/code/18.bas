Attribute VB_Name = "VbaHelper_GetWeekdayByDate"
Option Explicit

Function GetWeekdayByDate(ByVal d As Date) As String
    ' Gets weekday name by date
    GetWeekdayByDate = WeekdayName(Weekday(d, vbMonday), , vbMonday)
End Function