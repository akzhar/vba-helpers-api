Attribute VB_Name = "Helper74"
Option Explicit

Function RegExpReplace(ByVal text$, ByVal replacePattern$, ByVal replaceValue$)
    ' Replaces all occurrences of the substring in the original string
    Dim objRegExp As Object: Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = replacePattern
    objRegExp.Global = True
    RegExpReplace = objRegExp.Replace(text, replaceValue)
End Function