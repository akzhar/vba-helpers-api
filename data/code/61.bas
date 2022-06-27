Attribute VB_Name = "Helper61"
Option Explicit

Function RegExpTest(ByVal text$, ByVal pattern$) As Boolean
    ' ф-ция проверяет текст с помощью регулярного выражения
    ' возвращает True, если текст соответствует паттерну
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    RegExpTest = regExp.Test(text)
End Function