Attribute VB_Name = "Helper57"
Option Explicit

Function CopyRowFormulas(ByVal fromRow&, ByVal fromCol&, ByVal toCol&, ByVal targerRowFrom&, ByVal targerRowTo&)
    ' ф-ция копирует формулы из строки fromRow (колонки с fromCol по toCol) на строки с targerRowFrom по targerRowTo
    Range(Cells(fromRow, fromCol), Cells(fromRow, toCol)).Select
    Selection.Copy
    Range(Cells(targerRowFrom, fromCol), Cells(targerRowTo, toCol)).Select
    Selection.PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Function