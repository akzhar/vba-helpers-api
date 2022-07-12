Attribute VB_Name = "Helper29"
Option Explicit

Function GetFileName(ByVal filePath$) As String
    ' ф-ция возвращает имя файла с расширением (fileName.ext) из пути к нему
    Dim separator$: separator = Application.PathSeparator
    GetFileName = Split(filePath, separator)(UBound(Split(filePath, separator)))
End Function