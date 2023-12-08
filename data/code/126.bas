Attribute VB_Name = "VbaHelper_IsFolderExists"
Option Explicit

Function IsFolderExists(ByVal dirPath$) As Boolean
    ' Checks if specified folder exists
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    IsFolderExists = CBool(fso.FolderExists(dirPath))
    Set fso = Nothing
End Function