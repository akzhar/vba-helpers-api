Attribute VB_Name = "Helper7"
Option Explicit

Function FilterArr(ByRef arr(), ParamArray args()) As Variant
    ' ф-ция фильтрует многомерный массив arr, используя массив критериев фильтрации args
    ' формат критерия: "номер столбца" & "=" & "искомое значение", например, "3=маска текста"
    ' возвращает двумерный массив с подходящими строками из массива arr

    Dim i&, j&, filtersCount&, rowsCount&, rowNo&, mask$
    Dim filtersArr(), checksArr()

    On Error Resume Next
    FilterArr = Array()
    
    If UBound(args) = -1 Then
      Debug.Print "Error: filters required"
      Exit Function
    End if

    ReDim filters(0 To UBound(args) + 1, 1 To 2)
    Err.Clear
    
    i = UBound(arr, 2)

    If Err.Number > 0 Then
      Debug.Print "Error: two dimensional array required"
      Exit Function
    End if

    ' распознаем все параметры фильтрации
    For i = LBound(args) To UBound(args)
        mask = args(i)
        If Not IsMissing(mask) Then
            If mask Like "#*=*" Then
                filtersCount = filtersCount + 1
                filtersArr(filtersCount, 1) = Val(Split(mask, "=")(0)) ' столбец массива
                filtersArr(filtersCount, 2) = Split(mask, "=", 2)(1) ' маска для значения
            Else
                Debug.Print "Error: invalid filter '" & mask & "'"
            End If
        End If
    Next i

    If filtersCount = 0 Then
      Debug.Print "Error: all filters are empty"
      Exit Function
    End if
 
    ReDim checksArr(LBound(arr, 1) To UBound(arr, 1)) As Boolean

    ' проверяем все строки массива
    For i = LBound(arr, 1) To UBound(arr, 1)
        checksArr(i) = True
        ' перебираем все параметры фильтрации
        For j = 1 To filtersCount
            If Not (arr(i, filtersArr(j, 1)) Like filtersArr(j, 2)) Then
              checksArr(i) = False
              Exit For
            End if
        Next j
        ' увеличиваем счётчик подходящих строк на 1
        rowsCount = rowsCount - checksArr(i)
    Next i

    ' нет ни одной подходящей строки в массиве
    If rowsCount = 0 Then
      Debug.Print "There are no rows matched filter creterias"
      Exit Function
    End if
    
    ReDim filteredArr(0 To rowsCount - 1, LBound(arr, 2) To UBound(arr, 2))

    ' отбираем строки, которые были ранее помечены как подходящие
    For i = LBound(arr, 1) To UBound(arr, 1)
        If checksArr(i) Then 
            rowNo = rowNo + 1
            For j = LBound(arr, 2) To UBound(arr, 2)
                filteredArr(rowNo - 1, j) = arr(i, j)
            Next j
        End If
    Next i

    FilterArr = filteredArr

End Function