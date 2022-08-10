Attribute VB_Name = "Helper89"
Option Explicit

Function Filter2DArr(ByRef arr(), ByVal fnName$, ByVal elementPos&) As Variant()
    ' ф-ция фильтрует 2 мерный массив
    ' ф-ция с именем fnName будет вызвана с каждым эл-том массива arr (строкой)
    ' В кач-ве единственного параметра в ф-цию будет передан эл-т строки из столбца под номером elementPos
    ' если ф-ция fnName возвращает True, строка добавляется в результируюший массив
    
    Filter2DArr = Array()

    Dim i&: i = UBound(arr, 2)
    Dim checksArr() As Boolean: ReDim checksArr(LBound(arr, 1) To UBound(arr, 1))
    
    Dim rowsCount&, arrElement
    
    ' проверяем все строки массива
    For i = LBound(arr, 1) To UBound(arr, 1)
        checksArr(i) = False
        ' проверяем все строки массива
        arrElement = arr(i, elementPos)
        If Application.Run(fnName, arrElement) Then
            checksArr(i) = True
        End If
        ' увеличиваем счётчик подходящих строк на 1
        rowsCount = rowsCount - checksArr(i)
    Next i

    ' нет ни одной подходящей строки в массиве
    If rowsCount = 0 Then
        Debug.Print "There are no rows matched filter"
        Exit Function
    End If

    ReDim filteredArr(1 To rowsCount, LBound(arr, 2) To UBound(arr, 2))
    
    Dim rowNum&, j&
    
    ' отбираем строки, которые были ранее помечены как подходящие
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