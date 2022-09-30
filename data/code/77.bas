Attribute VB_Name = "Helper77"
Option Explicit

Function GETWORKDAYS(ByVal endDate As Date, ByVal startDate As Date) As Long
    ' Calculates the number of days between 2 dates minus the exception days
    Application.Volatile True
    Dim holidays: holidays = [Holidays].Value
    Dim daysDiff&: daysDiff = endDate - startDate
    Dim i&
    For i = LBound(holidays) To UBound(holidays)
        Dim holiday As Date: holiday = holidays(i, 1)
        If (startDate <= holiday And endDate >= holiday) Then
            ' Debug.Print ("Holiday: " & holidays(i, 1))
            daysDiff = daysDiff - 1
        End If
    Next i
    GETWORKDAYS = daysDiff
End Function