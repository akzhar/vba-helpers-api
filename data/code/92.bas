Attribute VB_Name = "Helper92"
Option Explicit

Function CreateWs(Optional ByVal tabName$, Optional ByVal tabColor& = -1, Optional ByRef afterWs As Worksheet, Optional ByVal needRecreate As Boolean = False) As Worksheet
    ' Creates a worksheet in the current workbook

    If tabName <> "" Then
        Dim wsExist As Boolean: wsExist = IsWsExists(ThisWorkbook, tabName) ' @(id 13)
        If wsExist Then
            If needRecreate Then
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets(tabName).Delete
                Application.DisplayAlerts = True
            Else
                Exit Function
            End If
        End If
    End If

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add( _
        After:=IIf( _
            afterWs Is Nothing, _
            ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), _
            afterWs _
        ) _
    )

    If tabName <> "" Then
        ws.name = tabName
    End If

    If tabColor <> -1 Then
        ws.Tab.color = tabColor
    End If
    
    Set CreateWs = ws
End Function