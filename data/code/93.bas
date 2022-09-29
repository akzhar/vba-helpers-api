Attribute VB_Name = "Helper93"
Option Explicit

Function MergeArrays(ByRef arr1(), ByRef arr2()) As Variant()
    ' ф-ция объединяет массивы и возвращает объединенный массив
    Dim arr3(): arr3 = arr1
    Dim i&
    For i = LBound(arr2) To UBound(arr2)
        Call AddToArr(arr3, arr2(i)) ' @(id 1)
    Next i
    MergeArrays = arr3
End Function