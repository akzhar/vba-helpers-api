Attribute VB_Name = "VbaHelper_GetColumnByHeader"
Option Explicit

Function GetColumnByHeader(ByRef ws As Worksheet, ByVal headerValue$, ByVal headerRow&) As Long
  ' Searches for the text in the specified row and returns the number of the column in which it was found
  
  Dim foundCol&: foundCol = 0
  Dim lastCol&: lastCol = GetLastColumn(ws, headerRow) ' @dependency: 65.bas
  
  Dim i&
  For i = 1 To lastCol
    If Trim(ws.Cells(headerRow, i).Value) = headerValue Then
        foundCol = i
        Exit For
    End If
  Next i

  GetColumnByHeader = Iif(foundCol = 0, -1, foundCol)

End Function
