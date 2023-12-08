Attribute VB_Name = "VbaHelper_GetWeekdayByNumber"
Option Explicit

Function GetWeekdayByNumber(ByVal n&) As String
    ' Gets weekday name by number
    Dim d As Dictionary: Set d = New Dictionary
    d.Add 1, "saturday": d.Add 2, "sunday": d.Add 3, "monday": d.Add 4, "tuesday"
    d.Add 5, "wednesday": d.Add 6, "thursday": d.Add 7, "friday"    
    GetWeekdayByNumber = d(n)    
End Function