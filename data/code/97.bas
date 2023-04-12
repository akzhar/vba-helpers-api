Attribute VB_Name = "VbaHelper_ClearDebugConsole"
Option Explicit

Function ClearDebugConsole()
    ' Clears VBE Immediate Window
    Application.VBE.Windows("Immediate").SetFocus
    If Application.VBE.ActiveWindow.Caption = "Immediate" And Application.VBE.ActiveWindow.Visible Then
        Application.SendKeys "^g ^a {DEL}"
    End If
End Function