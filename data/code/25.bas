Attribute VB_Name = "VbaHelper_GetFolderPath"
Option Explicit

Function GetFolderPath(ByVal titleMessage$, Optional ByVal defaultPath$ = "") As String
    ' Allows to select folder in dialog window
    
    Dim folderPath$: folderPath = ""
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