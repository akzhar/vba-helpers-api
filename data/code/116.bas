Attribute VB_Name = "VbaHelper_GetFileFromDialog"
Option Explicit

Function GetFileFromDialog(ByVal titleMessage$) As Workbook
    ' Open and returns an Excel file instance selected in dialog
    
    Dim folderPath$
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = titleMessage
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show <> -1 Then End
        folderPath = .SelectedItems(1)
    End With

    Set GetFileFromDialog = Application.Workbooks.Open(folderPath, False)
    
End Function