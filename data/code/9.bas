Attribute VB_Name = "VbaHelper_Sort2DArr"
Option Explicit

Function Sort2DArr(ByRef arr(), ByVal N&, Optional ByVal isDesc As Boolean = False) As Variant()
    ' Sorts 2-dim array by specified column N

    If N > UBound(arr, 1) Or N < LBound(arr, 1) Then
        MsgBox "Нет такого столбца в массиве", vbExclamation
        Exit Function
    End If

    Dim check As Boolean, i&, j&, condition As Boolean, tempArr()

    ReDim tempArr(UBound(arr, 2) + 1) As Variant

    Do Until check
        check = True
        For i = LBound(arr, 2) To UBound(arr, 2)
            condition = IIf( _
              isDesc, _
              arr(i, N) < arr(i + 1, N), _
              arr(i, N) > arr(i + 1, N) _
            )
            If condition Then
                For j = LBound(arr, 1) To UBound(arr, 1) - 1
                    tempArr(j) = arr(i, j)
                    arr(i, j) = arr(i + 1, j)
                    arr(i + 1, j) = tempArr(j)
                    check = False
                Next j
            End If
        Next i
    Loop

    Sort2DArr = arr
End Function