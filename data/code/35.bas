Attribute VB_Name = "Helper35"
Option Explicit

Function CopyColumnFormat(ByRef fromWs As Worksheet, ByVal fromCol&, ByRef targetWs As Worksheet, ByVal targetColFrom&, ByVal targetColTo&)
    ' Copies the format from the specified column and applies it to a range of columns (from ... to ...)
    fromWs.Columns(fromCol).Copy
    targetWs.Columns(targetColFrom).Resize(, targetColTo - 1).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function