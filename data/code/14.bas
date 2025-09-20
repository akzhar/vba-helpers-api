Attribute VB_Name = "VbaHelper_GetFileExtension"
Option Explicit

Function GetFileExtension(ByVal filePath$) As String
    ' Extracts the extension from the file path
    
    Dim lastDotPos&: lastDotPos = InStrRev(filePath, ".")
    Dim lastSlashPos&: lastSlashPos = InStrRev(filePath, "\")
    
    If lastDotPos > 0 And lastDotPos > lastSlashPos Then
        GetFileExtension = Mid(filePath, lastDotPos + 1)
    Else
        GetFileExtension = ""
    End If
End Function