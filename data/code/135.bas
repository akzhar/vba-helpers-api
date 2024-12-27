Attribute VB_Name = "VbaHelper_SplitArrToSubArrs"
Option Explicit

Function SplitArrToSubArrs(ByRef arr(), ByVal maxChunkSize&) As Variant()
    ' Splits an array into several subarrays of limited upper size
    Dim arrSize&: arrSize = UBound(arr) - LBound(arr) + 1    
    
    ' Calculate number of subarrays needed
    Dim numSubArrays&: numSubArrays = Int((arrSize + maxChunkSize - 1) / maxChunkSize)
    
    ' Initialize the subArrays array to hold the chunks
    Dim subArrays(): ReDim subArrays(0 To numSubArrays - 1)    
    
    Dim subArraySize&, i&, j&, tmp

    ' Split original array into chunks
    For i = 0 To numSubArrays - 1
        subArraySize = IIf(i = numSubArrays - 1, arrSize - i * maxChunkSize, maxChunkSize)
        tmp = subArrays(i)
        ReDim tmp(0 To subArraySize - 1)
        subArrays(i) = tmp        
        For j = 0 To subArraySize - 1
            subArrays(i)(j) = arr(i * maxChunkSize + j)
        Next j
    Next i    
   
    SplitArrToSubArrs = subArrays
  
End Function