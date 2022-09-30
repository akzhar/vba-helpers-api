Attribute VB_Name = "Helper86"
Option Explicit

Function CopyVBProject(ByRef srcWb As Workbook, ByRef wb As Workbook)
    ' Copies VBProject modules from one Excel file to another
    
    Dim separator$: separator = Application.PathSeparator
    Dim pathToSaveVba$: pathToSaveVba = Environ("Temp") & separator
    Dim modulePath$
    
    Dim objVbComp
    For Each objVbComp In srcWb.VBProject.VBComponents
      Select Case objVbComp.Type
         Case 1 'vbext_ct_StdModule
            modulePath = pathToSaveVba & separator & objVbComp.name & ".bas"
            objVbComp.Export modulePath
            wb.VBProject.VBComponents.Import modulePath
         Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            modulePath = pathToSaveVba & separator & objVbComp.name & ".cls"
            objVbComp.Export modulePath
            wb.VBProject.VBComponents.Import modulePath
         Case 3 'vbext_ct_MSForm
            modulePath = pathToSaveVba & separator & objVbComp.name & ".frm"
            objVbComp.Export modulePath
            wb.VBProject.VBComponents.Import modulePath
         Case Else
            modulePath = pathToSaveVba & separator & objVbComp.name
            objVbComp.Export modulePath
            wb.VBProject.VBComponents.Import modulePath
      End Select
    Next objVbComp
    
End Function