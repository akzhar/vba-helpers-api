Attribute VB_Name = "Helper24"
Option Explicit

Sub ExportVBProject()
    ' ф-ция экспортирует VBProject файлы из выбранного Excel файла в папку с именем macros

    Dim arr() as String: arr = GetFilePaths("Выбери файл с VBProject", "*.xlsm; *.xlsb") ' @(id 23)
    Dim pathToFile$: pathToFile = arr(0)
    Dim wb As Workbook: Set wb = Workbooks.Open(pathToFile)
    Dim pathToSaveVba As String: pathToSaveVba = wb.path
    
    Call CreateNewFolder(pathToSaveVba, "macros") ' @(id 22)
    
    Dim objVbComp
    For Each varItem In wb.VBProject.VBComponents
      Select Case objVbComp.Type
         Case 1 'vbext_ct_StdModule
            objVbComp.export pathToSaveVba & "\" & objVbComp.Name & ".bas"
         Case 2, 100 'vbext_ct_ClassModule, vbext_ct_Document
            objVbComp.export pathToSaveVba & "\" & objVbComp.Name & ".cls"
         Case 3 'vbext_ct_MSForm
            objVbComp.export pathToSaveVba & "\" & objVbComp.Name & ".frm"
         Case Else
            objVbComp.export pathToSaveVba & "\" & objVbComp.Name
      End Select
    Next objVbComp
    
    wb.Close False

    MsgBox "Готово", vbInformation
    
End Sub