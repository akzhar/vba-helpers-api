Attribute VB_Name = "Helper88"
Option Explicit

Function IsDateBetween(testDate As Date, startDate As Date, endDate As Date) As Boolean
    ' ф-ция проверяет находится ли дата в определенном диапазоне
    IsDateBetween = IIf(testDate >= startDate And testDate <= endDate, True, False)
End Function