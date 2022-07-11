Attribute VB_Name = "Helper14"
Option Explicit

Const CORRECT_PASSWORD$ = "qwerty"

Function VerifyPassword() As Boolean
    ' ф-ция запрашивает ввода пароля и проверяет его корректность
    Dim pass$: pass = InputBox("Введите пароль для продолжения:", "Пароль")
    VerifyPassword = IIf(pass = CORRECT_PASSWORD, True, False)
End Function