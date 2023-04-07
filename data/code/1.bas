Attribute VB_Name = "VbaHelper_AddToArr"
Option Explicit

Function AddToArr(ByRef arr(), ByVal element)
    ' Adds the specified element in 1 dim array
    If (Not arr) = -1 Then
        ReDim arr(0)
    Else
        ReDim Preserve arr(UBound(arr) + 1)
    End If
    arr(UBound(arr)) = element
End Function