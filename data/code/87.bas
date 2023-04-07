Attribute VB_Name = "VbaHelper_BreakLinks"
Option Explicit

Function BreakLinks(wb As Workbook)
    ' Breaks links in the specified workbook
    Dim linkSource: linkSource = wb.LinkSources(xlLinkTypeExcelLinks)
    On Error Resume Next
    Dim i&
    For i = 1 To UBound(linkSource)
      wb.BreakLink linkSource(i), xlLinkTypeExcelLinks
    Next i
    On Error GoTo 0
End Function