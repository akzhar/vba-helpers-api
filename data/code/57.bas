Attribute VB_Name = "Helper57"
Option Explicit

Function CopyRowFormulas(ByRef ws As Worksheet, ByVal fromCol&, ByVal toCol&, ByVal fromRow&, ByVal targerRowFrom&, ByVal targerRowTo&)
    ' Copies formulas from the specified row and applies them to a range of rows (from ... to ...)
    ws.Range(ws.Cells(fromRow, fromCol), ws.Cells(fromRow, toCol)).Copy
    ws.Range(ws.Cells(targerRowFrom, fromCol), ws.Cells(targerRowTo, toCol)).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Function