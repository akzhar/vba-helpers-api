Attribute VB_Name = "Helper57"
Option Explicit

Function CopyRowFormulas(ByRef ws As Worksheet, ByVal toCol&, ByVal targerRowFrom&, ByVal targerRowTo&, ByVal fromRow&, ByVal fromCol&)
    ' ф-ция копирует формулы из строки fromRow (колонки с fromCol по toCol) на строки с targerRowFrom по targerRowTo
    ws.Range(ws.Cells(fromRow, fromCol), ws.Cells(fromRow, toCol)).Copy
    ws.Range(ws.Cells(targerRowFrom, fromCol), ws.Cells(targerRowTo, toCol)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Function