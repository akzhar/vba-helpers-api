Attribute VB_Name = "Helper27"
Option Explicit

Function ReadTxtFile(ByVal filePath$, Optional ByVal encoding$ = "utf-8") As String
    ' Gets file's content in specified encoding from the specified txt file

    On Error Resume Next
    With CreateObject("ADODB.Stream")
        .Type = 2:
        If Len(encoding) Then .Charset = encoding
        .Open
        .LoadFromFile filePath
        ReadTxtFile = .ReadText
        .Close
    End With
    On Error GoTo 0
End Function
