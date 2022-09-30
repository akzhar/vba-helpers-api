Attribute VB_Name = "Helper12"
Option Explicit

Function IsWbOpen(ByVal wbName$) As Boolean
    ' Checks if specified workbook is open

    Dim wb As Workbook

    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0

    IsWbOpen = CBool(Not wb Is Nothing)
    
End Function