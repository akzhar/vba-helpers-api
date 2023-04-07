Attribute VB_Name = "VbaHelper_SplitToChars"
Option Explicit

Function SplitToChars(ByVal str$) As String()
    ' Splits the specified string into an array of characters / letters

    Dim arr() As String: ReDim arr(Len(str) - 1)
   
    Dim i&
    For i = 1 To Len(str)
        arr(i - 1) = Mid(str, i, 1)
    Next i
    
    SplitToChars = arr
End Function