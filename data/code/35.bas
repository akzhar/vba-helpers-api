Attribute VB_Name = "Helper35"
Option Explicit

Function CopyColumnFormat(ByRef fromWs As Worksheet, ByVal fromCol&, ByRef targetWs As Worksheet, ByVal targetColFrom&, ByVal targetColTo&)
    ' ф-ция копирует формат из столбца fromCol на диапазон столбов с targerColFrom по targerColTo
    fromWs.Columns(fromCol).Copy
    targetWs.Columns(targetColFrom).Resize(, targetColTo - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function