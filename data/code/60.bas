Attribute VB_Name = "Helper60"
Option Explicit

Function GetRegExpMatches(ByVal text$, ByVal pattern$) As Variant()
    ' ф-ция проверяет текст с помощью регулярного выражения
    ' возвращает массив со всеми совпадениями
    GetRegExpMatches = Array()
    Dim RegExp As Object: Set RegExp = CreateObject("VBScript.RegExp")
    RegExp.pattern = pattern
    RegExp.Global = True
    RegExp.MultiLine = True
    If RegExp.Test(text) Then
        Dim matchesColl As Object: Set matchesColl = RegExp.Execute(text)
        If matchesColl.Count <> 0 Then
            Dim i&, matches()
            For i = 0 To matchesColl.Count - 1
                Call AddToArr(matches, matchesColl(i)) ' @(id 1)
            Next i
            GetRegExpMatches = matches
        End If
    End If
End Function