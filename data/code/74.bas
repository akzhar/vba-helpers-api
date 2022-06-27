Attribute VB_Name = "Helper74"
Option Explicit

Function ReplaceAll(ByVal str$, ByVal replacePattern$, ByVal replaceValue$)
    ' ф-ция выполняет замену всех соответствий паттерну в строке
    Dim objRegExp As Object: Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = replacePattern
    objRegExp.Global = True
    ReplaceAll = objRegExp.Replace(str, replaceValue)
End Function