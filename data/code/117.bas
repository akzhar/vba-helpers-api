Attribute VB_Name = "VbaHelper_CountString"
Option Explicit

Function CountString(ByVal text$, ByVal str$) As Long
    ' Counts how many times a string contains a sub string 
    CountString = (Len(text) - Len(Replace(text, str, ""))) / Len(str)
End Function