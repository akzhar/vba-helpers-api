Attribute VB_Name = "Helper115"
Option Explicit

Function GetFileExtension(ByVal filePath$) As String
    ' Extracts the extension from the file path
    GetFileExtension = Right(filePath, Len(filePath) - InStr(1, filePath, ".") + 1)
End Function