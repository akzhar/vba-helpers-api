Attribute VB_Name = "Helper19"
Option Explicit

Function GetWeekNum(ByVal d As Date) As Long
    ' Gets week number in year by date
    GetWeekNum = DatePart("ww", d, vbMonday)
End Function