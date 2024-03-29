Attribute VB_Name = "VbaHelper_Save2TxtFile"
Option Explicit

Function Save2TxtFile(ByVal text$, ByVal filePath$, ByVal fileName$, Optional ByVal encoding = "utf-8") As String
    ' Writes content in txt file in specified encoding and save the file in specified location 

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

    Set FO = Nothing

    Save2TxtFile = fullPath
    
End Function