Attribute VB_Name = "VbaHelper_NestedArrTo2DArr"
Option Explicit

Private Function NestedArrTo2DArr(ByVal arr) As Variant()
    ' Transforms nested array (array of 1-dim arrays) into 2-dim array
    
    Dim i&, j&, res()

    Dim rowsCount&: rowsCount = UBound(arr) - LBound(arr) + 1
    Dim colsCount&: colsCount = UBound(arr(LBound(arr))) - LBound(arr(LBound(arr))) + 1
    
    ReDim res(1 To rowsCount, 1 To colsCount)
    
    For i = LBound(arr) To UBound(arr)
        For j = LBound(arr(i)) To UBound(arr(i))
            res(i - LBound(arr) + 1, j - LBound(arr(i)) + 1) = arr(i)(j)
        Next j
    Next i
    
    NestedArrTo2DArr = res
    
End Function
