Attribute VB_Name = "VbaHelper_Number2Text"
Option Explicit

Function Number2Text(ByVal N As Long) As String
    ' Returns number as text
    
    Dim arrNumbers(19) As String, arrTens(9) As String, arrHundreds(9) As String
    
    arrNumbers(0) = "ноль"
    arrNumbers(1) = "один"
    arrNumbers(2) = "два"
    arrNumbers(3) = "три"
    arrNumbers(4) = "четыре"
    arrNumbers(5) = "пять"
    arrNumbers(6) = "шесть"
    arrNumbers(7) = "семь"
    arrNumbers(8) = "восемь"
    arrNumbers(9) = "девять"
    arrNumbers(10) = "десять"
    arrNumbers(11) = "одиннадцать"
    arrNumbers(12) = "двенадцать"
    arrNumbers(13) = "тринадцать"
    arrNumbers(14) = "четырнадцать"
    arrNumbers(15) = "пятнадцать"
    arrNumbers(16) = "шестнадцать"
    arrNumbers(17) = "семнадцать"
    arrNumbers(18) = "восемнадцать"
    arrNumbers(19) = "девятнадцать"
    
    arrTens(2) = "двадцать"
    arrTens(3) = "тридцать"
    arrTens(4) = "сорок"
    arrTens(5) = "пятьдесят"
    arrTens(6) = "шестьдесят"
    arrTens(7) = "семьдесят"
    arrTens(8) = "восемьдесят"
    arrTens(9) = "девяносто"
    
    arrHundreds(1) = "сто"
    arrHundreds(2) = "двести"
    arrHundreds(3) = "триста"
    arrHundreds(4) = "четыреста"
    arrHundreds(5) = "пятьсот"
    arrHundreds(6) = "шестьсот"
    arrHundreds(7) = "семьсот"
    arrHundreds(8) = "восемьсот"
    arrHundreds(9) = "девятьсот"
    
    Dim numberAsText$
    
    If N >= 1000000 Then
        Number2Text = "слишком большое число"
        Exit Function
    End If
    
    If N >= 1000 Then
        numberAsText = Number2Text(N \ 1000) & " тысяч"
        N = N Mod 1000
    End If
    
    If N >= 100 Then
        numberAsText = numberAsText & " " & arrHundreds(N \ 100)
        N = N Mod 100
    End If
    
    If N >= 20 Then
        numberAsText = numberAsText & " " & arrTens(N \ 10)
        N = N Mod 10
    End If
    
    If N > 0 Then
        If N < 20 Then
            numberAsText = numberAsText & " " & arrNumbers(N)
        Else
            numberAsText = numberAsText & " " & arrNumbers(N Mod 10)
        End If
    End If
    
    Number2Text = Trim(numberAsText)
    
End Function