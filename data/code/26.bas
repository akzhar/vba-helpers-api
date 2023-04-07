Attribute VB_Name = "VbaHelper_ReadTxtFile"
Option Explicit

Function ReadTxtFile(ByVal filePath$) As String
    ' Gets file's content from the specified txt file

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim srcFile As Object: Set srcFile = fso.OpenTextFile(filePath, 1)
    ' read whole file
    Dim data$: data = srcFile.ReadAll
    ' read the line by line
    'Dim line$
    'While Not srcFile.AtEndOfStream
    '    line = srcFile.ReadLine
        ' do something with the line...
    'Wend
    srcFile.Close

    Set fso = Nothing
    Set srcFile = Nothing

    ReadTxtFile = data
End Function