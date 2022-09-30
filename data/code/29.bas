Attribute VB_Name = "Helper29"
Option Explicit

Function GetFileName(ByVal filePath$) As String
    ' Extracts the name with the extension from the file path
    Dim separator$: separator = Application.PathSeparator
    GetFileName = Split(filePath, separator)(UBound(Split(filePath, separator)))
End Function