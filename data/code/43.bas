Attribute VB_Name = "VbaHelper_ShowProcessing"
Option Explicit

Function ShowProcessing(ByVal flag As Boolean)
    ' Shows operation execution message in Excel status bar
    Select Case flag
      Case True
        Application.DisplayStatusBar = True
        Application.Cursor = xlWait
        Application.StatusBar = "Operation in progress. Please wait..."
      Case False
        Application.StatusBar = False
        Application.Cursor = xlDefault
    End Select
End Function
