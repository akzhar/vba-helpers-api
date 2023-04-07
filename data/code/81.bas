Attribute VB_Name = "VbaHelper_Collection2Array"
Option Explicit

Function Collection2Array(ByRef coll As Object) As Variant()
    ' Converts collection to array
    Dim i&, arr()
    For i = 1 To coll.Count
        Call AddToArr(arr, coll(i))  ' @dependency: 1.bas
    Next i
    Collection2Array = arr
End Function