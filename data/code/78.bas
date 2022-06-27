Attribute VB_Name = "Helper78"
Option Explicit

Function ValidateInput(ByVal textRequest$, ByVal checkPattern$, ByVal textPattern$, Optional ByVal textWarning$, Optional ByVal defaultInput$) As String
    ' ф-ция проверяет ввод пользователя и возвращает вводимые данные

    Dim inputData$
    
    If textWarning <> "" Then textWarning = vbLf & vbLf & "Внимание!" & vbLf & textWarning

SelectData:

    inputData = Trim(InputBox(textRequest & vbLf & vbLf & "Соблюдайте формат ввода: " & textPattern & textWarning, "Введите данные", defaultInput))
    
    ' проверка, что ввод не пустой
    If inputData = "" Then
        If MsgBox("Поле ввода пустое." & vbLf & vbLf & "Повторить ввод?", vbYesNo, "Ошибка") = vbYes Then
            GoTo SelectData
        Else
            Exit Function
        End If
    End If
    
    ' проверка, что ввод соответствует формату ввода
    If Not inputData Like checkPattern Then
        If MsgBox("Введенное значение не соответствует формату ввода: " & textPattern & "." & vbLf & vbLf & "Повторить ввод?", vbYesNo, "Ошибка") = vbYes Then
            GoTo SelectData
        Else
            Exit Function
        End If
    End If

    ValidateInput = inputData
End Function