Attribute VB_Name = "VbaHelper_GetTimeStamp"
Option Explicit

Private Type SystemTime
    wYear          As Integer
    wMonth         As Integer
    wDayOfWeek     As Integer
    wDay           As Integer
    wHour          As Integer
    wMinute        As Integer
    wSecond        As Integer
    wMilliseconds  As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystem As SystemTime)

Function GetTimeStamp() As String
    ' Gets current timestamp (yyyy-mm-dd hh:mm:ss:mss)
    Dim d As Date: d = Now()
    Dim t As SystemTime
    GetLocalTime t
    GetTimeStamp = Format(d, "yyyy") & "-" & Format(d, "MM") & "-" & Format(d, "dd") & " " & Format(t.wHour, "00") & ":" & Format(t.wMinute, "00") & ":" & Format(t.wSecond, "00") & ":" & Format(t.wMilliseconds, "000")
End Function