Attribute VB_Name = "VbaHelper_CombineArrays"
Option Explicit

Function CombineArrays(ByRef arr1(), ByRef arr2()) As Variant()
    ' Combines 2 arrays
    Dim arr(): arr = arr1
    Dim i&
    For i = LBound(arr2) To UBound(arr2)
        Call AddToArr(arr, arr2(i)) ' @dependency: 1.bas
    Next i
    CombineArrays = arr
End Function