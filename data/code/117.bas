Attribute VB_Name = "VbaHelper_CountSubstringInString"
Option Explicit

Function CountSubstringInString(ByVal str$, ByVal subStr$) As Long
    ' Counts how many times a string contains a sub string 
    CountSubstringInString = Len(str) - Len(Replace(str, subStr, ""))
End Function