Attribute VB_Name = "VbaHelper_AddToArr"
Option Explicit

Function AddToArr(ByRef arr(), ByVal element, Optional ByVal startsFrom& = 0)
    ' Adds the specified element in 1 dim array
    If (Not arr) = -1 Then
        ReDim arr(startsFrom To startsFrom)
    Else
        ReDim Preserve arr(startsFrom To UBound(arr) + 1)
    End If
    arr(UBound(arr)) = element
End Function