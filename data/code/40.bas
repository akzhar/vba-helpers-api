Attribute VB_Name = "VbaHelper_SendHttpQuery"
Option Explicit

Private HttpCodes As Object

Private Function InitHttpCodes()
    Set HttpCodes = CreateObject("Scripting.Dictionary")
    HttpCodes("200") = "OK"
    HttpCodes("400") = "Bad Request"
    HttpCodes("404") = "Not Found"
    HttpCodes("500") = "Internal Server Error"
End Function

Function SendHttpQuery(ByVal url$, Optional ByVal method$ = "GET", Optional ByVal contentType$ = "text/plain", Optional ByVal reqBody$) As Variant
    ' Executes HTTP query

    Call InitHttpCodes
    
    Dim req As Object: Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
          
    With req
        .Open method, url, False
        .setRequestHeader "Content-Type", contentType & "; charset=UTF-8"
        .send reqBody
    End With
    
    If req.Status <> "200" Then
        MsgBox "Server response is not OK." _
        & vbLf & vbLf & req.Status & ": " & HttpCodes(Cstr(req.Status)), vbExclamation
        Exit Function
    End If
    
    If contentType = "application/json" Then
        Set SendHttpQuery = ParseJson(req.responseText) ' @dependency: 42.bas
    Else
        SendHttpQuery = req.responseText
    End If
            
End Function