Attribute VB_Name = "Helper20"
Option Explicit

Function GetDateByDayNum(ByVal dayNum&, ByVal year$, Optional ByVal dateFormat$ = "dd.mm.yyyy") As String
    ' ф-ция возвращает дату в заданном формате по порядковому номеру дня в году
    GetDateByDayNum = Format(DateSerial(year, 1, dayNum), dateFormat)
End Function