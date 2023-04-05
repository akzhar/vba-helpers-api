Attribute VB_Name = "Helper29"
Option Explicit

Function GetFileName(ByVal filePath$) As String
    ' Extracts the name with the extension from the file path
    GetFileName = Split(filePath, Application.PathSeparator)(UBound(Split(filePath, Application.PathSeparator)))
End Function