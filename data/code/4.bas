Attribute VB_Name = "VbaHelper_IsInArray"
Option Explicit

Function IsInArray(ByRef arr(), ByVal element) As Boolean
    ' Checks if array contains the specified element
    Dim i&

    IsInArray = False
    
    If (Not arr) = -1 Then Exit Function
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = element Then
            IsInArray = True
            Exit Function
        End If
    Next i

End Function
