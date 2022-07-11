Attribute VB_Name = "Helper25"
Option Explicit

Function GetFolderPath(ByVal titleMessage$, Optional ByVal defaultPath$ = "") As String
    ' ф-ция открывает окно для выбора папки
    ' возвращает путь к выбранной папке
    
    Dim folderPath$
    Dim dialog As FileDialog: Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialog
        .Title = titleMessage
        .AllowMultiSelect = False
        .InitialFileName = IIf(defaultPath = "", ThisWorkbook.path, defaultPath)
        If .Show <> -1 Then
            GoTo CancelHandler
        End If
        folderPath = .SelectedItems(1)
    End With
    
CancelHandler:
    
    Set dialog = Nothing
        
    If Len(folderPath) > 0 Then
        GetFolderPath = folderPath
    End If
    
End Function