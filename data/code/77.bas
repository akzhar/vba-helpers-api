Attribute VB_Name = "Helper77"
Option Explicit

Function GETWORKDAYS(ByVal endDate As Date, ByVal startDate As Date) As Long
    ' ф-ция возвращает кол-во дней между 2-мя датами за вычетом праздничных дней
    ' даты праздников указываются в Диспетчере имен --> создать диапазон с именем "Праздники"
    Application.Volatile True
    Dim holidays: holidays = [Праздники].Value
    Dim daysDiff&: daysDiff = endDate - startDate
    Dim i&
    For i = LBound(holidays) To UBound(holidays)
        Dim holiday As Date: holiday = holidays(i, 1)
        If (startDate <= holiday And endDate >= holiday) Then
            ' Debug.Print ("Праздник: " & holidays(i, 1))
            daysDiff = daysDiff - 1
        End If
    Next i
    GETWORKDAYS = daysDiff
End Function