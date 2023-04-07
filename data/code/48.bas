Attribute VB_Name = "VbaHelper_FixNumbers"
Option Explicit

Function FixNumbers(ByRef rng As Range)
    ' Fixes number stored as text error
    
    Dim cell As Range
    Call TurnUpdatesOn(False) ' @dependency: 51.bas
    For Each cell In rng
        If Len(cell) > 0 Then
            cell.Value2 = CDbl(Replace(cell.Value2, ",", Application.DecimalSeparator, 1, -1, vbBinaryCompare))
        End If
    Next cell
    Call TurnUpdatesOn(True) ' @dependency: 51.bas
End Function