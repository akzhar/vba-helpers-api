Attribute VB_Name = "Helper54"
Option Explicit

Function ClearFilters(ByRef ws As Worksheet)
    ' Clears all autofilters in worksheet
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
End Function