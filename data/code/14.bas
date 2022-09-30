Attribute VB_Name = "Helper14"
Option Explicit

Const CORRECT_PASSWORD$ = "qwerty"

Function VerifyPassword() As Boolean
    ' Prompts you to enter a password and checks its correctness
    Dim pass$: pass = InputBox("To continue run macros please type the password:", "Need password")
    VerifyPassword = IIf(pass = CORRECT_PASSWORD, True, False)
End Function