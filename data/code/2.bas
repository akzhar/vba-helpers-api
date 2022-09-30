Attribute VB_Name = "Helper2"
Option Explicit

Function GetArrLength(ByRef arr()) As Long
    ' Gets length of the specified array
    If IsEmpty(arr) Then
        GetArrLength = 0
    Else
        GetArrLength = UBound(arr) - LBound(arr) + 1
    End If
End Function