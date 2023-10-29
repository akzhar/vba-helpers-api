Attribute VB_Name = "VbaHelper_TurnUpdatesOn"
Option Explicit

Function TurnUpdatesOn(ByVal flag As Boolean)
    ' Turns off or on the updates of Excel interface (screen updates, formulas calculation, events, status bar, alerts)
    Application.ScreenUpdating = flag
    Application.Calculation = IIf(flag = True, xlCalculationAutomatic, xlCalculationManual)
    Application.EnableEvents = flag
    Application.DisplayStatusBar = flag
    Application.DisplayAlerts = flag
    Application.AskToUpdateLinks = flag
End Function 