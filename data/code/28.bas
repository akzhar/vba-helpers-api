Attribute VB_Name = "Helper28"
Option Explicit

Function SaveToTxtFile(ByVal text$, ByVal filePath$, ByVal fileName$, Optional ByVal encoding = "utf-8")
    ' ф-ция сохраняет текст в txt файл
    Dim FO As Object: Set FO = CreateObject("ADODB.Stream")
    Dim separator$: separator = IIf(Right(filePath, 1) <> Application.PathSeparator, Application.PathSeparator, "")
    Dim fullPath$: fullPath = filePath & separator & fileName
    With FO
        .Type = 2 ' specify stream type (text/string data)
        .Charset = encoding ' specify charset for the source text data
        .Open ' open the stream
        .WriteText text ' write binary data to the object
        .SaveToFile fullPath, 2 ' save binary data to disk
    End With
    Debug.Print ("File '" & fullPath & "' was created")
    Set FO = Nothing
End Function