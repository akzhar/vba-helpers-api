Attribute VB_Name = "Helper13"
Option Explicit

Function IsWsExists(ByRef wb As Workbook, ByVal wsName$) As Boolean
    ' ф-ция проверяет наличие листа в Excel книге по его имени
    IsWsExists = False
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If wsName = ws.name Then
            IsWsExists = True
            Exit Function
        End If
    Next ws
End Function