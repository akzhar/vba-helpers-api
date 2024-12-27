Attribute VB_Name = "VbaHelper_CopyWs"
Option Explicit

Function CopyWs(ByRef wsToCopy As Worksheet, ByVal afterWs As Worksheet, ByVal copiedWsName$)
    ' Creates a copy of the sheet
    wsToCopy.Copy After:=afterWs
    If IsWsExists(ThisWorkbook, copiedWsName) Then ' @dependency: 13.bas
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(copiedWsName).Delete
        Application.DisplayAlerts = True
    End If
    wsToCopy.name = copiedWsName
    ThisWorkbook.ActiveSheet.Tab.Color = vbYellow
End Function