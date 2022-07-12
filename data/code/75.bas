Attribute VB_Name = "Helper75"
Option Explicit

Dim startTime&
Dim endTime&
Dim executionTime$

Function RunTimer()
    startTime = Timer()
End Function

Function StopTimer() As String
    endTime = Timer()
    executionTime = Format((endTime - startTime) / 86400, "hh:mm:ss")
    stopTimer = executionTime
End Function