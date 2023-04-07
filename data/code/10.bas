Attribute VB_Name = "VbaHelper_GetUniqueArr"
Option Explicit

Function GetUniqueArr(ByRef arr()) As Variant()
    ' Get 1-dim array without duplicate values
    
    Dim uniqueArr(), isDuplicate As Boolean
    
    Dim i&
    For i = LBound(arr) To UBound(arr)
        isDuplicate = IsInArray(uniqueArr, arr(i)) ' @dependency: 4.bas
        If Not isDuplicate Then
            Call AddToArr(uniqueArr, arr(i))  ' @dependency: 1.bas
        End If
    Next i

    GetUniqueArr = uniqueArr
End Function