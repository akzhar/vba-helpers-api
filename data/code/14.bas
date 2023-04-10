Attribute VB_Name = "VbaHelper_GetFileExtension"
Option Explicit

Function GetFileExtension(ByVal filePath$) As String
    ' Extracts the extension from the file path
    GetFileExtension = Right(filePath, Len(filePath) - InStr(1, filePath, ".") + 1)
End Function