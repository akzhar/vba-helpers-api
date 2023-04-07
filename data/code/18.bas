Attribute VB_Name = "VbaHelper_GetWeekday"
Option Explicit

Function GetWeekday(ByVal d As Date) As String
    ' Gets weekday name by date
    GetWeekday = WeekdayName(Weekday(d, vbMonday))
End Function