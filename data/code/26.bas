Attribute VB_Name = "Helper26"
Option Explicit

Function ReadTxtFile(ByVal filePath$) As String
    ' ф-ция считывает txt файл и возвращает его содержимое

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