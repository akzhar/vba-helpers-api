Attribute VB_Name = "Helper87"
Option Explicit

Function BreakLinks(wb As Workbook)
    ' ф-ция удаляет связи (Edit Links) из переданной Excel книги
    Dim linkSource: linkSource = wb.LinkSources(xlLinkTypeExcelLinks)
    On Error Resume Next
    Dim i&
    For i = 1 To UBound(linkSource)
      wb.BreakLink linkSource(i), xlLinkTypeExcelLinks
    Next i
    On Error GoTo 0
End Function