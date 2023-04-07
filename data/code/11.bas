Attribute VB_Name = "VbaHelper_GetIndexOf"
Option Explicit

Function GetIndexOf(ByRef arr(), ByVal element, Optional ByVal elementPos = Null) As Long
    ' Get index of specified element in 1-dim array
    
    GetIndexOf = -1

    Dim i&, arrElement As Variant
    
    For i = LBound(arr) To UBound(arr)
        If IsNull(elementPos) Then
            arrElement = arr(i)
        Else
            arrElement = arr(i, elementPos)
        End If
        If element = arrElement Then
            GetIndexOf = i
            Exit Function
        End If
    Next i
    
End Function