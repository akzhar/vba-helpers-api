Attribute VB_Name = "Helper61"
Option Explicit

Function RegExpTest(ByVal text$, ByVal pattern$) As Boolean
    ' Checks if text matches the specified regular expression pattern
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    RegExpTest = regExp.Test(text)
End Function