Attribute VB_Name = "VbaHelper_VerifyPassword"
Option Explicit

Private Const CORRECT_PASSWORD$ = "qwerty"

Function VerifyPassword() As Boolean
    ' Prompts you to enter a password and checks its correctness
    Dim pass$: pass = InputBox("To continue run the macros please enter the password:", "Need password")
    VerifyPassword = IIf(pass = CORRECT_PASSWORD, True, False)
End Function