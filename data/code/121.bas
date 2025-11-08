Attribute VB_Name = "VbaHelper_CopyWs"
Option Explicit

Function CopyWs(ByRef wsToCopy As Worksheet, ByRef afterWs As Worksheet, ByVal copiedWsName$, Optional ByVal copiedtabColor& = -1)
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

    ThisWorkbook.ActiveSheet.Name = copiedWsName

    If copiedtabColor <> -1 Then
        ThisWorkbook.ActiveSheet.Tab.color = copiedtabColor
    Else
        ThisWorkbook.ActiveSheet.Tab.color = vbYellow
    End If
    
End Function