Attribute VB_Name = "Helper43"
Option Explicit

Function ShowProcessing(ByVal flag As Boolean)
    ' ф-ция показывает / скрывает сообщение о выполнении операции в статус баре Excel
    If flag = True Then
        Application.DisplayStatusBar = True
        Application.Cursor = xlWait
        Application.StatusBar = "Операция выполняется. Пожалуйста, подождите..."
    End If
    If flag = False Then
        Application.StatusBar = False
        Application.Cursor = xlDefault
    End If
End Function
