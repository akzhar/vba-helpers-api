Attribute VB_Name = "Helper73"
Option Explicit

Function Rng2String(ByRef rng As Range, ByVal separator$) As String
    ' Concatenates all values from the range into a text string separated by the specified separator
    Dim arr() As String: arr = Rng2Array(rng) ' @(id 3)
    Rng2String = Join(arr, separator)
End Function