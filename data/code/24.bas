Attribute VB_Name = "Helper24"
Option Explicit

Function ExportVBProject()
    ' Exports VBProject modules from the selected Excel workbook

    Const FOLDER_NAME$ = "macros"
    Dim arr() As String: arr = GetFilePaths("Выбери файл с VBProject", "*.xlsm; *.xlsb") ' @dependency: 23.bas
    Dim pathToFile$: pathToFile = arr(0)
    Dim wb As Workbook: Set wb = Workbooks.Open(pathToFile)
    
    Dim pathToSaveVba As String: pathToSaveVba = wb.path
    
    Call CreateFolder(pathToSaveVba, FOLDER_NAME) ' @dependency: 22.bas
    
    pathToSaveVba = pathToSaveVba & Application.PathSeparator & FOLDER_NAME
    
    Dim ext$, objVbComp
    
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
      objVbComp.Export pathToSaveVba & Application.PathSeparator & objVbComp.Name & ext
    Next objVbComp
    
    wb.Close False

    MsgBox "VBProject files export completed", vbInformation
    
End Function