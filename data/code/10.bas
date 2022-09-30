Attribute VB_Name = "Helper10"
Option Explicit

Function GetUniqueArr(ByRef arr()) As Variant()
    ' Get 1-dim array without duplicate values
    
    Dim uniqueArr(): ReDim uniqueArr(0)
    Dim isDuplicate As Boolean, arrIndex&, newArrIndex&: newArrIndex = 0
    
    For arrIndex = LBound(arr) To UBound(arr)
        isDuplicate = IsInArray(uniqueArr, arr(arrIndex)) ' @(id 4)
        If Not isDuplicate Then
            ReDim Preserve uniqueArr(newArrIndex)
            uniqueArr(newArrIndex) = arr(arrIndex)
            newArrIndex = newArrIndex + 1
        End If
    Next arrIndex

    GetUniqueArr = uniqueArr
End Function