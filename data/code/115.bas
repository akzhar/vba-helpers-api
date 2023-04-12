Attribute VB_Name = "VbaHelper_RepeatStrNTimes"
Option Explicit

Function RepeatStrNTimes(ByVal str$, ByVal n&) As String
    ' Concatenates string several times
    
    If n < 0 Then Exit Function
    Dim arr() As String: ReDim arr(n)
    RepeatStrNTimes = Join(arr, str)
    
End Function