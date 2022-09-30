Attribute VB_Name = "Helper75"
Option Explicit

Dim startTime&
Dim endTime&
Dim executionTime$

Function RunTimer(ByVal flag As Boolean) As String
    ' Counts the execution time of the functions / procedures
    If flag Then
        startTime = Timer()
    Else
        endTime = Timer()
        executionTime = Format((endTime - startTime) / 86400, "hh:mm:ss")
        RunTimer = executionTime
    End if
End Function