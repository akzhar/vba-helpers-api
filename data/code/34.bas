Attribute VB_Name = "VbaHelper_CopyRowFormats"
Option Explicit

Function CopyRowFormats(ByRef fromWs As Worksheet, ByVal fromRow&, ByRef targetWs As Worksheet, ByVal targerRowFrom&, ByVal targerRowTo&)
    ' Copies the format from the specified row and applies it to a range of rows (from ... to ...)
    fromWs.Rows(fromRow & ":" & fromRow).Copy
    targetWs.Rows(targerRowFrom & ":" & targerRowTo).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Function