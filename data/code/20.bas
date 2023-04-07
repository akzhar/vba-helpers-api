Attribute VB_Name = "VbaHelper_GetDateByDayNum"
Option Explicit

Function GetDateByDayNum(ByVal dayNum&, ByVal yearNum&, Optional ByVal dateFormat$ = "dd.mm.yyyy") As String
    ' Gets text representation of the date in the specified format by number of the day in year
    GetDateByDayNum = Format(DateSerial(yearNum, 1, dayNum), dateFormat)
End Function