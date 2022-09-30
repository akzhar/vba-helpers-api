Attribute VB_Name = "Helper40"
Option Explicit

Const PROXY_URL$ = "xxx.proxy.su"
Dim HttpCodes As Object

Function initHttpCodes()
    Set HttpCodes = CreateObject("Scripting.Dictionary")
    HttpCodes("200") = "OK"
    HttpCodes("400") = "BAD REQUEST"
    HttpCodes("404") = "NOT FOUND"
    HttpCodes("500") = "SERVER ERROR"
End Function

Function HttpQuery(ByVal url$, Optional ByVal method$ = "GET", Optional ByVal contentType$ = "text/plain", Optional ByVal reqBody$) As Variant
    ' Executes HTTP query

    Call initHttpCodes
    
    Dim req As Object: Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
          
    With req
        .Open method, url, False
        If PROXY_URL <> "" Then
            .setProxy 2, PROXY_URL, ""
        End If
        .setRequestHeader "Content-Type", contentType & "; charset=UTF-8"
        .send reqBody
    End With
    
    If req.Status <> "200" Then
        MsgBox "Server response is not OK." _
        & vbLf & vbLf & "Response status: " & req.Status & HttpCodes(req.Status), vbExclamation
        Exit Function
    End If
    
    If contentType = "application/json" Then
        Set HttpQuery = JsonConverter.ParseJson(req.responseText) ' @(id 42)
    Else
        HttpQuery = req.responseText
    End If
            
End Function