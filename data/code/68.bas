Attribute VB_Name = "Helper68"
Option Explicit

Function GetRowByHeader(ByRef ws As Worksheet, ByVal headerValue$, ByVal headerCol&) As Long
  ' Searches for the text in the specified column and returns the number of the row in which it was found

  Dim foundRow&: foundRow = 0
  Dim lastRow&: lastRow = GetLastRow(ws, headerCol) ' @(id 64)
  Dim i&

  For i = 1 To lastRow
    If Trim(ws.Cells(i, headerCol).Value) = headerValue Then
        foundRow = i
        Exit For
    End If
  Next i

  If foundRow = 0 Then
    GetRowByHeader = -1
    Exit Function
  End If

  GetRowByHeader = foundRow
End Function