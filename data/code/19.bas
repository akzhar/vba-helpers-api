Attribute VB_Name = "Helper19"
Option Explicit

Function GetWeekNum(ByVal d As Date) As Long
    ' ф-ция возвращает номер недели по дате
    GetWeekNum = DatePart("ww", d, vbMonday)
End Function