Attribute VB_Name = "Helper71"
Option Explicit

Function SplitToChars(ByVal str$) As String()
    ' ф-ция разбивает строку на массив символов

    Dim arr() As String: ReDim arr(Len(str) - 1)
   
    Dim i&
    For i = 1 To Len(str)
        arr(i - 1) = Mid(str, i, 1)
    Next i
    
    SplitToChars = arr
End Function