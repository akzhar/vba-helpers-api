Attribute VB_Name = "Helper77"
Option Explicit

Function GETWORKDAYS(ByVal startDate As Date, ByVal endDate As Date) As Long
    ' Calculates the number of days between 2 dates minus the exception days
    Application.Volatile True
    Dim holidays: holidays = [Holidays].Value
    Dim daysDiff&: daysDiff = endDate - startDate
    Dim i&
    For i = LBound(holidays) To UBound(holidays)
        Dim holiday As Date: holiday = holidays(i, 1)
        If (holiday >= startDate And holiday <= endDate) Then
            daysDiff = daysDiff - 1
        End If
    Next i
    GETWORKDAYS = daysDiff
End Function