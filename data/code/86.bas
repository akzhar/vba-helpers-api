Attribute VB_Name = "VbaHelper_CopyVBProject"
Option Explicit

Function CopyVBProject(ByRef srcWb As Workbook, ByRef wb As Workbook)
    ' Copies VBProject modules from one Excel file to another

    Dim ext$, objVbComp

    For Each objVbComp In srcWb.VBProject.VBComponents
      Select Case objVbComp.Type
         Case 1
            ext = ".bas" 'vbext_ct_StdModule
         Case 2, 100
            ext = ".cls" 'vbext_ct_ClassModule, vbext_ct_Document
         Case 3
            ext =".frm" 'vbext_ct_MSForm
         Case Else
            ext = ""
      End Select
      Dim modulePath$: modulePath = Environ("Temp") & Application.PathSeparator & objVbComp.name & ext
      objVbComp.Export modulePath
      wb.VBProject.VBComponents.Import modulePath
    Next objVbComp
    
End Function