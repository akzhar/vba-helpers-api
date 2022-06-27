Attribute VB_Name = "Helper35"
Option Explicit

Function CopyColFormatting(ByRef fromWs As Worksheet, ByVal fromCol&, ByRef targetWs As Worksheet, ByVal targerColFrom&, ByVal targerColTo&)
    ' ф-ция копирует формат из столбца fromCol на диапазон столбов с targerColFrom по targerColTo
    fromWs.Columns(fromCol & ":" & fromCol).Copy
    targetWs.Columns(targetColFrom & ":" & targerColTo).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function