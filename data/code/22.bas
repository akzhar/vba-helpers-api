Attribute VB_Name = "Helper22"
Option Explicit

Function CreateFolder(ByVal dirPath$, ByVal dirName$) As Boolean
    ' Creates a folder in specified location with specified name

    CreateFolder = False

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path$: path = dirPath & Application.PathSeparator & dirName
    Dim dirExist As Boolean: dirExist = CBool(Dir(path, vbDirectory) <> "")
    
    If Not dirExist Then
        MkDir (path)
        CreateFolder = True
    End If
End Function