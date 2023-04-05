Attribute VB_Name = "Helper113"
Option Explicit

Function IsModuleExists(ByVal vbModuleName$, ByRef wb As Workbook) As Boolean
    ' Checks if VBProject contains specified module
    
    Dim ext$, objVbComp
    
    IsModuleExists = False
    
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
      IsModuleExists = CBool((objVbComp.Name & ext) = vbModuleName)
      If IsModuleExists Then Exit Function
    Next objVbComp
    
End Function