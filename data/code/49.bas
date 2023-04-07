Attribute VB_Name = "VbaHelper_Round"
Option Explicit

Function Round(ByVal strNumber$, ByVal numDigits&) As String
    ' Rounds the specified float number to N digits after the decimal separator

    Dim delimiter$: delimiter = Application.DecimalSeparator
    Dim delimiterPos As Integer: delimiterPos = InStr(1, strNumber, delimiter, vbTextCompare)

    If delimiterPos = 0 Then
        Round = strNumber
    Else
        Dim leftPart&: leftPart = CLng(Left(strNumber, delimiterPos))
        Dim decimalPart$: decimalPart = Right(strNumber, Len(strNumber) - delimiterPos)
        Dim ePos&: ePos = InStr(1, LCase(decimalPart), "e+")
        If ePos > 0 Then
            decimalPart = Left(decimalPart, ePos - 1)
        End If
        
        Dim isNegative As Boolean: isNegative = InStr(1, leftPart, "-", vbTextCompare) <> 0
        
        Dim digits() As String: digits = Split(StrConv(CStr(decimalPart), 64), Chr(0))
        
        Dim i As Integer: i = UBound(digits) - 1

        While (i > IIf(numDigits = 0, numDigits, numDigits - 1))
            Dim digit As Integer: digit = CInt(digits(i))
            Dim prevDigit As Integer: prevDigit = CInt(digits(i - 1))
            Dim j As Integer: j = i - 2
            While (prevDigit = 9 And j >= 0)
                prevDigit = CInt(digits(j))
                j = j - 1
            Wend
            If digit >= 5 And prevDigit <> 9 Then
                digits(i) = 0
                digits(i - 1) = prevDigit + 1
            End If
            i = i - 1
        Wend
        
        If numDigits = 0 Then
            If digits(0) >= 5 Then
                Round = CDbl(IIf(isNegative, leftPart - 1, leftPart + 1))
            Else
                Round = CDbl(leftPart)
            End If
        Else
            Round = CDbl(Left(CStr(leftPart & delimiter & Join(digits, "")), delimiterPos + numDigits))
        End If

    End If
End Function
