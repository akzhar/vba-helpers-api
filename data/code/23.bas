Attribute VB_Name = "Helper23"
Option Explicit

Function GetFilePaths(ByVal titleMessage$, Optional ByVal extensionFilters$ = "", Optional ByVal defaultPath$ = ThisWorkbook.Path, Optional ByVal allowMulti As Boolean = False) As String()
    ' ф-ция открывает окно для выбора файлов
    ' возвращает массив с путями к выбранным файлам
    ' если ничего не выбрано вернет пустой массив

    Dim pathsArr() As String, i As Integer

    Dim dialog As Object: Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)

    With dialog
        .Title = titleMessage
        .AllowMultiSelect = allowMulti
        .InitialFileName = defaultPath
        If extensionFilters <> "" Then
            .Filters.Clear
            .Filters.Add Description:="Only allowed extensions", Extensions:=extensionFilters
        End If
        .Show
    End With

    For i = 1 To dialog.SelectedItems.Count
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