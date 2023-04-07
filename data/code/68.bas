Attribute VB_Name = "VbaHelper_GetRowByHeader"
Option Explicit

Function GetRowByHeader(ByRef ws As Worksheet, ByVal headerValue$, ByVal headerCol&) As Long
  ' Searches for the text in the specified column and returns the number of the row in which it was found

  Dim foundRow&: foundRow = 0
  Dim lastRow&: lastRow = GetLastRow(ws, headerCol) ' @dependency: 64.bas
  Dim i&

  For i = 1 To lastRow
    If Trim(ws.Cells(i, headerCol).Value) = headerValue Then
        foundRow = i
        Exit For
    End If
  Next i

  GetRowByHeader = Iif(foundRow = 0, -1, foundRow)

End Function