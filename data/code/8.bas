Attribute VB_Name = "VbaHelper_SortArr"
Option Explicit

Function SortArr(ByRef arr(), Optional ByVal isDesc As Boolean = False) As Variant()
    ' Sorts 1-dim array
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