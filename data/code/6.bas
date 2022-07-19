Attribute VB_Name = "Helper6"
Option Explicit

Function FilterArr(ByRef arr(), ByVal fnName$, Optional ByVal elementPos&) As Variant()
    ' ф-ция фильтрует 1 или 2 мерный массив
    ' возвращает 1 мерный массив со всеми вхождениями element в arr
    ' ф-ция с именем fnName будет вызвана с каждым эл-том массива arr в кач-ве единственного параметра
    ' если ф-ция fnName возвращает True, element добавляется в результируюший массив
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