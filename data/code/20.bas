Attribute VB_Name = "Helper20"
Option Explicit

Function GetDateByDayNum(ByVal dayNum&, ByVal year&, Optional ByVal dateFormat$ = "dd.mm.yyyy") As String
    ' Gets text representation of the date in the specified format by number of the day in year
    GetDateByDayNum = Format(DateSerial(year, 1, dayNum), dateFormat)
End Function