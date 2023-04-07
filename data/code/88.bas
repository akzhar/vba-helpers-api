Attribute VB_Name = "VbaHelper_IsDateBetween"
Option Explicit

Function IsDateBetween(testDate As Date, startDate As Date, endDate As Date) As Boolean
    ' Checks if the specified date is in a date range
    IsDateBetween = IIf(testDate >= startDate And testDate <= endDate, True, False)
End Function