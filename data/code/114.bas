Attribute VB_Name = "VbaHelper_DownloadFile"
Option Explicit

Function DownloadFile(ByVal url$) As String
    ' Downloads the file from specified url and saves it into the Temp folder
    
    Dim WinHttpReq As Object: Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", url, False
    WinHttpReq.send

    Dim fileName$: fileName = Split(url, "/")(UBound(Split(url, "/")))
    
    Dim pathToSave$: pathToSave = Environ("temp") & Application.PathSeparator & fileName

    If WinHttpReq.Status = 200 Then
        Dim oStream As Object: Set oStream = CreateObject("ADODB.Stream")
        With oStream
            .Open
            .Type = 1
            .Write WinHttpReq.responseBody
            .SaveToFile pathToSave, 2 ' 1 = no overwrite, 2 = overwrite
            .Close
        End With
        DownloadFile = pathToSave
    Else
        Debug.Print (WinHttpReq.Status & ": " & WinHttpReq.statusText)
        DownloadFile = ""
    End If

    Set WinHttpReq = Nothing

End Function