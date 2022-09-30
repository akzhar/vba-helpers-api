Attribute VB_Name = "Helper89"
Option Explicit

Function Filter2DArr(ByRef arr(), ByVal fnName$, ByVal elementPos&) As Variant()
    ' Filters 2-dim array using callback checker-function
    
    Filter2DArr = Array()

    Dim i&: i = UBound(arr, 2)
    Dim checksArr() As Boolean: ReDim checksArr(LBound(arr, 1) To UBound(arr, 1))
    
    Dim rowsCount&, arrElement
    
    ' check all the rows
    For i = LBound(arr, 1) To UBound(arr, 1)
        checksArr(i) = False
        arrElement = arr(i, elementPos)
        If Application.Run(fnName, arrElement) Then
            checksArr(i) = True
        End If
        rowsCount = rowsCount - checksArr(i)
    Next i

    If rowsCount = 0 Then
        Debug.Print "There are no rows matched filter"
        Exit Function
    End If

    ReDim filteredArr(1 To rowsCount, LBound(arr, 2) To UBound(arr, 2))
    
    Dim rowNum&, j&
    
    ' collect all the rows which passed filtering
    For i = LBound(arr, 1) To UBound(arr, 1)
        If checksArr(i) Then
            rowNum = rowNum + 1
            For j = LBound(arr, 2) To UBound(arr, 2)
                filteredArr(rowNum, j) = arr(i, j)
            Next j
        End If
    Next i

    Filter2DArr = filteredArr
End Function