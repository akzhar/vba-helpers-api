Attribute VB_Name = "VbaHelper_GetFirstRegExpMatch"
Option Explicit

Function GetFirstRegExpMatch(ByVal text$, ByVal pattern$) As String
    ' Gets only the 1st regular expression match from the text
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    If regExp.Test(text) Then
        GetFirstRegExpMatch = regExp.Execute(text)(0).Value
    End If
End Function