Attribute VB_Name = "VbaHelper_CopySheet"
Option Explicit

Function CopySheet(ByRef wsToCopy As Worksheet, ByVal afterWs As Worksheet, ByVal copiedWsName$)
    ' Creates a copy of the sheet
    If Utils.IsWsExists(ThisWorkbook, copiedWsName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(copiedWsName).Delete
        Application.DisplayAlerts = True
    End If
    wsToCopy.name = copiedWsName
    wsToCopy.Copy After:=afterWs
    ThisWorkbook.Sheets(copiedWsName).Tab.Color = vbYellow
End Function