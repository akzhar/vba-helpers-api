Attribute VB_Name = "Helper34"
Option Explicit

Function CopyRowFormatting(ByRef fromWs As Worksheet, ByVal fromRow&, ByRef targetWs As Worksheet, ByVal targerRowFrom&, ByVal targerRowTo&)
    ' ф-ция копирует формат из строки fromRow на диапазон строк с targerRowFrom по targerRowTo
    fromWs.Rows(fromRow & ":" & fromRow).Copy
    targetWs.Rows(targerRowFrom & ":" & targerRowTo).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function