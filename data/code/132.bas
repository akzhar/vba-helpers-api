Attribute VB_Name = "VbaHelper_GetDateFromText"
Option Explicit

Function GetDateFromText(ByVal str$) As Date
    ' Gets date from a text string
    On Error Resume Next
    GetDateFromText = DateValue(str)
    On Error GoTo 0
    If GetDateFromText <> 0 Then Exit Function
    Dim tmp() As String
    ' @dependency: 61.bas
    Select Case True
        Case RegExpTest(str, "^\d{1,2}\.\d{2}\.\d{4}$")
            tmp = Split(str, ".")
            GetDateFromText = DateSerial(CInt(tmp(2)), CInt(tmp(1)), CInt(tmp(0)))
        Case RegExpTest(str, "^\d{1,2}/\d{2}/\d{4}$")
            tmp = Split(str, "/")
            GetDateFromText = DateSerial(CInt(tmp(2)), CInt(tmp(0)), CInt(tmp(1)))
        Case RegExpTest(str, "^\d{4}-\d{2}\-\d{2}$")
            tmp = Split(str, "-")
            GetDateFromText = DateSerial(CInt(tmp(0)), CInt(tmp(1)), CInt(tmp(2)))
    End Select
End Function
