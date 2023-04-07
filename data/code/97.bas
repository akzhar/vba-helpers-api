Attribute VB_Name = "VbaHelper_ClearDebugConsole"
Option Explicit

Function ClearDebugConsole()
    ' Clears VBE Immediate Window
    Dim i&
    For i = 0 To 100
        Debug.Print ""
    Next i
End Function