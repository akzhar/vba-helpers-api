Attribute VB_Name = "VbaHelper_GetArrLength"
Option Explicit

Function GetArrLength(ByRef arr()) As Long
    ' Gets length of the specified array
    On Error Resume Next
    Dim IsInitialised As Boolean: IsInitialised = IsNumeric(UBound(arr))
    On Error GoTo 0
    If IsEmpty(arr) Or Not IsInitialised(arr) Then
        GetArrLength = 0
    Else
        GetArrLength = UBound(arr) - LBound(arr) + 1
    End If
End Function