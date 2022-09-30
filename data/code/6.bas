Attribute VB_Name = "Helper6"
Option Explicit

Function FilterArr(ByRef arr(), ByVal fnName$, Optional ByVal elementPos&) As Variant()
    ' Filters 1 or 2 dim array using callback checker-function
    Dim i&, arrElement, filteredArr()
    
    filteredArr = Array()

    For i = LBound(arr) To UBound(arr)
        If elementPos = 0 Then
            arrElement = arr(i)
        Else
            arrElement = arr(i, elementPos)
        End If
        If Application.Run(fnName, arrElement) Then
            ReDim Preserve filteredArr(UBound(filteredArr) + 1)
            filteredArr(UBound(filteredArr)) = arrElement
        End If
    Next i

    FilterArr = filteredArr
End Function