Attribute VB_Name = "VbaHelper_IsWsExists"
Option Explicit

Function IsWsExists(ByRef wb As Workbook, ByVal wsName$) As Boolean
    ' Checks if specified worsheet exists in the workbook

    IsWsExists = False

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If wsName = ws.name Then
            IsWsExists = True
            Exit Function
        End If
    Next ws
    
End Function