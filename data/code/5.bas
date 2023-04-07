Attribute VB_Name = "VbaHelper_FilterArr"
Option Explicit

Function FilterArr(ByRef arr(), ByVal element, Optional ByVal elementPos = Null) As Variant()
    ' Filters 1 or 2 dim array
    Dim i&, arrElement, filteredArr()
    
    filteredArr = Array()

    For i = LBound(arr) To UBound(arr)
        If IsNull(elementPos) Then
            arrElement = arr(i)
        Else
            arrElement = arr(i, elementPos)
        End If
        If element = arrElement Then
            ReDim Preserve filteredArr(UBound(filteredArr) + 1)
            filteredArr(UBound(filteredArr)) = arrElement
        End If
    Next i

    FilterArr = filteredArr
End Function