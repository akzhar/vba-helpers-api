Attribute VB_Name = "Helper93"
Option Explicit

Function CombineArrays(ByRef arr1(), ByRef arr2()) As Variant()
    ' Combines 2 arrays
    Dim arr3(): arr3 = arr1
    Dim i&
    For i = LBound(arr2) To UBound(arr2)
        Call AddToArr(arr3, arr2(i)) ' @(id 1)
    Next i
    CombineArrays = arr3
End Function