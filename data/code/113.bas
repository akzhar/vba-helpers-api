Attribute VB_Name = "VbaHelper_IsVBModuleExists"
Option Explicit

Function IsVBModuleExists(ByVal vbModuleName$, ByRef wb As Workbook) As Boolean
    ' Checks if VBProject contains specified module
    
    Dim ext$, objVbComp
    
    IsVBModuleExists = False
    
    For Each objVbComp In wb.VBProject.VBComponents
      Select Case objVbComp.Type
         Case 1
            ext = ".bas" 'vbext_ct_StdModule
         Case 2, 100
            ext = ".cls" 'vbext_ct_ClassModule, vbext_ct_Document
         Case 3
            ext = ".frm" 'vbext_ct_MSForm
         Case Else
            ext = ""
      End Select
      IsVBModuleExists = CBool((objVbComp.Name & ext) = vbModuleName)
      If IsVBModuleExists Then Exit Function
    Next objVbComp
    
End Function