Attribute VB_Name = "Helper9"
Option Explicit

Function SortArr(ByRef arr(), ByVal N&, Optional ByVal isDesc As Boolean = True) As Variant()
    ' ф-ция сортирует переданный 2 мерный массив по столбцу N
    ' desc: по убыванию, от большего к меньшему (<)
    ' asc: по возрастанию, от меньшего к большему (>)

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

    SortArr = arr
End Function