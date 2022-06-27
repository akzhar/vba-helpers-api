Attribute VB_Name = "Helper48"
Option Explicit

Function ConvertNumbersStoredAsText(ByRef rng As Range)
    ' ф-ция исправляет ошибку number stored as text, преобразую каждую ячейку в диапазоне в число
    Dim cell As Range
    Call TurnUpdates(False) ' @(id 51)
    For Each cell In rng
        If Len(cell) > 0 Then
            cell.Value2 = CDbl(Replace(cell.Value2, ",", Application.DecimalSeparator, 1, -1, vbBinaryCompare))
        End If
    Next cell
    Call TurnUpdates(True) ' @(id 51)
End Function