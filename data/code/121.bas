Attribute VB_Name = "VbaHelper_CopyWs"
Option Explicit

Function CopyWs(ByRef wsToCopy As Worksheet, ByRef afterWs As Worksheet, ByVal copiedWsName$)
    ' Creates a copy of the sheet
    If IsWsExists(ThisWorkbook, copiedWsName) Then ' @dependency: 13.bas
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(copiedWsName).Delete
        Application.DisplayAlerts = True
    End If
    With wsToCopy
        '.Visible = True
        .Copy After:=afterWs
        '.Visible = False
    End With
    ThisWorkbook.ActiveSheet.name = copiedWsName
    ThisWorkbook.ActiveSheet.Tab.Color = vbYellow
End Function