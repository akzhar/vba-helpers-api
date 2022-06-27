Attribute VB_Name = "Helper62"
Option Explicit

Function GetFirstRegExpMatch(ByVal text$, ByVal pattern$) As String
    ' ф-ция проверяет текст с помощью регулярного выражения
    ' возвращает первое соответствие текста паттерну
    Dim regExp As Object: Set regExp = CreateObject("VBScript.RegExp")
    regExp.Pattern = pattern
    regExp.Global = True
    regExp.MultiLine = True
    If regExp.Test(text) Then
        GetFirstRegExpMatch = regExp.Execute(text)(0).Value
    End If
End Function