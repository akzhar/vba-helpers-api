Attribute VB_Name = "Helper67"
Option Explicit

Function GetColumnByHeader(ByRef ws As Worksheet, ByVal headerValue$, ByVal headerRow&) As Long
  ' ф-ция возвращает номер столбца, в котором был найден заголовок headerValue в строке headerRow
  
  Dim foundCol&: foundCol = 0
  Dim lastCol&: lastCol = GetLastColumn(ws, headerRow) ' @(id 65)
  
  Dim i&
  For i = 1 To lastCol
    If Trim(ws.Cells(headerRow, i).Value) = headerValue Then
        foundCol = i
        Exit For
    End If
  Next i

  If foundCol = 0 Then
    GetColumnByHeader = -1
    Exit Function
  End If

  GetColumnByHeader = foundCol
End Function
