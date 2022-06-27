Attribute VB_Name = "Helper68"
Option Explicit

Function GetRowByHeader(ByRef ws As Worksheet, ByVal headerValue$, ByVal headerCol&) As Long
  ' ф-ция возвращает номер строки, в которой был найден заголовок headerValue в столбце headerCol

  Dim foundRow&: foundRow = 0
  Dim lastRow&: lastRow = utils.GetLastRow(ws, headerCol)
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