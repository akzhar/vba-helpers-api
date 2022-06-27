Attribute VB_Name = "Helper19"
Option Explicit

Function GetWeekNumByDate(ByVal d$) As Long
    ' ф-ция возвращает номер недели по дате
    GetWeekNumByDate = DatePart("ww", CDate(d), vbMonday)
End Function