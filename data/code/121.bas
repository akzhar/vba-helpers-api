Attribute VB_Name = "VbaHelper_CopyWs"
Option Explicit

Function CopyWs(ByRef wsToCopy As Worksheet, ByRef afterWs As Worksheet, ByVal copiedWsName$, Optional ByVal copiedtabColor& = -1)
    ' Creates a copy of the sheet

    Dim wsName$: wsName = Left(copiedWsName, 31)

    If IsWsExists(ThisWorkbook, wsName) Then ' @dependency: 13.bas
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(wsName).Delete
        Application.DisplayAlerts = True
    End If

    With wsToCopy
        '.Visible = True
        .Copy After:=afterWs
        '.Visible = False
    End With

    ThisWorkbook.ActiveSheet.Name = wsName

    If copiedtabColor <> -1 Then
        ThisWorkbook.ActiveSheet.Tab.color = copiedtabColor
    Else
        ThisWorkbook.ActiveSheet.Tab.color = vbYellow
    End If
    
End Function