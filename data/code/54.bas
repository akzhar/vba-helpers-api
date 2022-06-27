Attribute VB_Name = "Helper54"
Option Explicit

Function RemoveFilters(ByRef ws As Worksheet)
    ' ф-ция снимает установленные фильтры с листа
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
End Function