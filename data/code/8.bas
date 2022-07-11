Attribute VB_Name = "Helper8"
Option Explicit

Function SortArr(ByRef arr(), Optional ByVal isDesc As Boolean = True) As Variant
    ' ф-ция сортирует пузырьком 1 мерный массив arr
    ' desc: по убыванию, от большего к меньшему (<)
    ' asc: по возрастанию, от меньшего к большему (>)
    Dim i&, j&, condition As Boolean, temp
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            condition = Iif(isDesc, arr(i) < arr(j), arr(i) > arr(j))
            If condition Then
                temp = arr(j)
                arr(j) = arr(i)
                arr(i) = temp
            End If
        Next j
    Next i

    SortArr = arr
End Function