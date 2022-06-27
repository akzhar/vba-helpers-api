Attribute VB_Name = "Helper14"
Option Explicit

Function VerifyPassword() As Boolean
    ' ф-ция запрашивает ввода пароля и проверяет его корректность
    Dim input$: input = InputBox("Введите пароль для продолжения", "Пароль")
    VerifyPassword = Iif(Cbool(input = CORRECT_PASSWORD), True, False)
End Function