Attribute VB_Name = "VbaHelper_FilterArr2"
Option Explicit

Function FilterArr2(ByRef arr(), ByVal fnName$, Optional ByVal elementPos = Null) As Variant()
    ' Filters 1 or 2 dim array using callback checker-function
    Dim i&, arrElement, filteredArr()
    
    filteredArr = Array()

    For i = LBound(arr) To UBound(arr)
        If IsNull(elementPos) Then
            arrElement = arr(i)
        Else
            arrElement = arr(i, elementPos)
        End If
        If Application.Run(fnName, arrElement) Then
            ReDim Preserve filteredArr(UBound(filteredArr) + 1)
            filteredArr(UBound(filteredArr)) = arrElement
        End If
    Next i

    FilterArr2 = filteredArr
End Function