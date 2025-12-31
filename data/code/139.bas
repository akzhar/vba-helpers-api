Attribute VB_Name = "VbaHelper_Capitalize"
Option Explicit

Function Capitalize(ByVal str$)
    ' Converts the first character of a string to uppercase (capital letter)
    Capitalize = UCase(Left(str, 1)) & LCase(Mid(str, 2))
End Function

