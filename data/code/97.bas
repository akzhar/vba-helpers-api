Attribute VB_Name = "Helper97"
Option Explicit

Function ClearDebugConsole()
    ' Clears VBE Immediate Window
    Dim i&
    For i = 0 To 100
        Debug.Print ""
    Next i
End Function