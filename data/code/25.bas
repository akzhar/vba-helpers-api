Attribute VB_Name = "Helper25"
Option Explicit

Function GetFolderPath(ByVal titleMessage$, Optional ByVal defaultPath$ = ThisWorkbook.Path) As String
    ' ф-ция открывает окно для выбора папки
    ' возвращает путь к выбранной папке
    
    Dim folderPath$
    Dim dialog As FileDialog: Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    With dialog
        .Title = titleMessage
        .AllowMultiSelect = False
        .InitialFileName = defaultPath
        If .Show <> -1 Then
            GoTo NextCode
        End If
        folderPath = .SelectedItems(1)
    End With
    
NextCode:
    
    Set dialog = Nothing
        
    If Len(folderPath) > 0 Then
        GetFolderPath = folderPath
    End If
    
End Function