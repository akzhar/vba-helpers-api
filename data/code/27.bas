Attribute VB_Name = "Helper27"
Option Explicit

Function ReadTxtFile(ByVal filePath$, Optional ByVal encoding$ = "windows-1251") As String
    ' ф-ция считывает txt файл в указанной кодировке и возвращает его содержимое
    On Error Resume Next:
    With CreateObject("ADODB.Stream")
        .Type = 2:
        If Len(encoding) Then .Charset = encoding
        .Open
        .LoadFromFile filePath
        ReadTxtFile = .ReadText
        .Close
    End With
End Function
