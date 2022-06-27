Attribute VB_Name = "Helper60"
Option Explicit

Function RegExpExtract(ByVal text$, ByVal pattern$) As Object
    ' ф-ция проверяет текст с помощью регулярного выражения
    ' возвращает объект-коллекцию со всеми совпадениями
    Dim RegExp As Object: Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.Pattern = pattern
    RegExp.Global = True
    regExp.MultiLine = True
    If RegExp.Test(text) Then
        Set RegExpExtract = RegExp.Execute(text)
    End If
End Function