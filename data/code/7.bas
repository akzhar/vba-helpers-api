Attribute VB_Name = "Helper7"
Option Explicit

Function Filter2DArr(ByRef arr(), ParamArray args()) As Variant()
    ' ф-ция фильтрует многомерный массив arr, используя массив критериев фильтрации args
    ' формат критерия: "номер столбца" & "=" & "искомое значение", например, "3=маска текста"
    ' возвращает двумерный массив с подходящими строками из массива arr
    Filter2DArr = Array()
    
    On Error Resume Next

    If UBound(args) = -1 Then
        Debug.Print "Error: filters required"
        Exit Function
    End If

    ReDim Filters(0 To UBound(args) + 1, 1 To 2)
    Err.Clear

    Dim i&: i = UBound(arr, 2)

    If Err.Number > 0 Then
        Debug.Print "Error: two dimensional array required"
        Exit Function
    End If

    Dim arg$, mask$, col&
    Dim filtersCount&
    Dim filtersArr(): ReDim filtersArr(UBound(args), 1)

    ' распознаем все параметры фильтрации
    For i = LBound(args) To UBound(args)
        arg = args(i)
        If Not IsMissing(arg) Then
            If arg Like "#*=*" Then
                filtersCount = filtersCount + 1
                col = Val(Split(arg, "=")(0)) ' столбец со значением
                mask = Split(arg, "=", 2)(1) ' маска для значения
                filtersArr(i, 0) = col
                filtersArr(i, 1) = mask
            Else
                Debug.Print "Error: invalid filter '" & arg & "'"
            End If
        End If
    Next i

    If filtersCount = 0 Then
        Debug.Print "Error: all filters are empty"
        Exit Function
    End If

    Dim rowsCount&, j&
    Dim checksArr() As Boolean: ReDim checksArr(LBound(arr, 1) To UBound(arr, 1))

    ' проверяем все строки массива
    For i = LBound(arr, 1) To UBound(arr, 1)
        checksArr(i) = True
        ' перебираем все параметры фильтрации
        For j = 1 To filtersCount
            col = filtersArr(j - 1, 0)
            mask = filtersArr(j - 1, 1)
            If Not (arr(i, col) Like mask) Then
                checksArr(i) = False
                Exit For
          End If
      Next j
      ' увеличиваем счётчик подходящих строк на 1
      rowsCount = rowsCount - checksArr(i)
    Next i

    ' нет ни одной подходящей строки в массиве
    If rowsCount = 0 Then
        Debug.Print "There are no rows matched filter creterias"
        Exit Function
    End If

    Dim rowNum&
    ReDim filteredArr(0 To rowsCount - 1, LBound(arr, 2) To UBound(arr, 2))

    ' отбираем строки, которые были ранее помечены как подходящие
    For i = LBound(arr, 1) To UBound(arr, 1)
        If checksArr(i) Then
            rowNum = rowNum + 1
            For j = LBound(arr, 2) To UBound(arr, 2)
                filteredArr(rowNum - 1, j) = arr(i, j)
            Next j
        End If
    Next i

    Filter2DArr = filteredArr
End Function