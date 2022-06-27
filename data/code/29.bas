Attribute VB_Name = "Helper29"
Option Explicit

Function GetFileNameFromPath(ByVal filePath$) As String
    ' ф-ция возвращает имя файла с расширением (fileName.ext) из пути к нему
    GetFileNameFromPath = Split(filePath, "\")(UBound(Split(filePath, "\")))
End Function