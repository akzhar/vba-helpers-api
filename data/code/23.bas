Attribute VB_Name = "VbaHelper_GetFilePaths"
Option Explicit

Function GetFilePaths(ByVal titleMessage$, Optional ByVal extensionFilters$ = "", Optional ByVal defaultPath$ = "", Optional ByVal allowMulti As Boolean = False) As String()
    ' Allows to select files in dialog window

    Dim pathsArr() As String, i&

    Dim dialog As Object: Set dialog = Application.FileDialog(msoFileDialogFilePicker)

    With dialog
        .Title = titleMessage
        .AllowMultiSelect = allowMulti
        .InitialFileName = IIf(defaultPath = "", ThisWorkbook.path, defaultPath)
        If extensionFilters <> "" Then
            .Filters.Clear
            .Filters.Add Description:="Only allowed extensions", Extensions:=extensionFilters
        End If
        .Show
    End With

    For i = 1 To dialog.SelectedItems.count
        ReDim Preserve pathsArr(i - 1)
        pathsArr(i - 1) = dialog.SelectedItems(i)
    Next i

    Set dialog = Nothing
    
    If Join(pathsArr) = Empty Then
        Exit Function
    Else
        GetFilePaths = pathsArr
    End If
    
End Function