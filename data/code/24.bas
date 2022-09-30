Attribute VB_Name = "Helper24"
Option Explicit

Function ExportVBProject()
    ' Exports VBProject modules from the selected Excel workbook

    Const FOLDER_NAME$ = "macros"
    Dim arr() As String: arr = GetFilePaths("Выбери файл с VBProject", "*.xlsm; *.xlsb") ' @(id 23)
    Dim pathToFile$: pathToFile = arr(0)
    Dim wb As Workbook: Set wb = Workbooks.Open(pathToFile)
    Dim pathToSaveVba As String: pathToSaveVba = wb.path
    Dim separator$: separator = Application.PathSeparator
    
    Call CreateFolder(pathToSaveVba, FOLDER_NAME) ' @(id 22)
    
    pathToSaveVba = pathToSaveVba & separator & FOLDER_NAME
    
    Dim objVbComp
    For Each objVbComp In wb.VBProject.VBComponents
      Select Case objVbComp.Type
         Case 1 'vbext_ct_StdModule
            objVbComp.Export pathToSaveVba & separator & objVbComp.Name & ".bas"
         Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            objVbComp.Export pathToSaveVba & separator & objVbComp.Name & ".cls"
         Case 3 'vbext_ct_MSForm
            objVbComp.Export pathToSaveVba & separator & objVbComp.Name & ".frm"
         Case Else
            objVbComp.Export pathToSaveVba & separator & objVbComp.Name
      End Select
    Next objVbComp
    
    wb.Close False

    MsgBox "VBProject files export completed", vbInformation
    
End Function