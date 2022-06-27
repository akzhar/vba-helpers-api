Attribute VB_Name = "Helper51"
Option Explicit

Function TurnUpdatesOn(ByVal flag as Boolean)
    ' ф-ция вкл / выкл обновление экрана, пересчет формул, события, статус-бар, алерты
    Application.ScreenUpdating = flag
    Application.Calculation = IIf(flag = True, xlCalculationAutomatic, xlCalculationManual)
    Application.EnableEvents = flag
    Application.DisplayStatusBar = flag
    Application.DisplayAlerts = flag
End Function 